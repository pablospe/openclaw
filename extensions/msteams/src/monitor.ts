import type { Request, Response } from "express";
import {
  DEFAULT_WEBHOOK_MAX_BODY_BYTES,
  isDangerousNameMatchingEnabled,
  keepHttpServerTaskAlive,
  mergeAllowlist,
  summarizeMapping,
  type OpenClawConfig,
  type RuntimeEnv,
} from "../runtime-api.js";
import { createMSTeamsConversationStoreFs } from "./conversation-store-fs.js";
import type { MSTeamsConversationStore } from "./conversation-store.js";
import { formatUnknownError } from "./errors.js";
import { runMSTeamsFileConsentInvokeHandler } from "./file-consent-invoke.js";
import { registerMSTeamsHandlers, type MSTeamsActivityHandler } from "./monitor-handler.js";
import type { MSTeamsMessageHandlerDeps } from "./monitor-handler.types.js";
import {
  createMSTeamsPollStoreFs,
  extractMSTeamsPollVote,
  type MSTeamsPollStore,
} from "./polls.js";
import {
  resolveMSTeamsChannelAllowlist,
  resolveMSTeamsUserAllowlist,
} from "./resolve-allowlist.js";
import { getMSTeamsRuntime } from "./runtime.js";
import type { MSTeamsTurnContext } from "./sdk-types.js";
import {
  createMSTeamsExpressAdapter,
  createMSTeamsTokenProvider,
  loadMSTeamsSdkWithAuth,
  type MSTeamsApp,
  type MSTeamsCardActionResponse,
} from "./sdk.js";
import { runMSTeamsSigninInvokeHandler } from "./signin-invoke.js";
import { createMSTeamsSsoTokenStoreFs } from "./sso-token-store.js";
import type { MSTeamsSsoDeps } from "./sso.js";
import { resolveMSTeamsCredentials } from "./token.js";
import { applyMSTeamsWebhookTimeouts } from "./webhook-timeouts.js";

type MonitorMSTeamsOpts = {
  cfg: OpenClawConfig;
  runtime?: RuntimeEnv;
  abortSignal?: AbortSignal;
  conversationStore?: MSTeamsConversationStore;
  pollStore?: MSTeamsPollStore;
};

type MonitorMSTeamsResult = {
  app: unknown;
  shutdown: () => Promise<void>;
};

export async function monitorMSTeamsProvider(
  opts: MonitorMSTeamsOpts,
): Promise<MonitorMSTeamsResult> {
  const core = getMSTeamsRuntime();
  const log = core.logging.getChildLogger({ name: "msteams" });
  let cfg = opts.cfg;
  let msteamsCfg = cfg.channels?.msteams;
  if (!msteamsCfg?.enabled) {
    log.debug?.("msteams provider disabled");
    return { app: null, shutdown: async () => {} };
  }

  const creds = resolveMSTeamsCredentials(msteamsCfg);
  if (!creds) {
    log.error("msteams credentials not configured");
    return { app: null, shutdown: async () => {} };
  }
  const appId = creds.appId; // Extract for use in closures

  const runtime: RuntimeEnv = opts.runtime ?? {
    log: console.log,
    error: console.error,
    exit: (code: number): never => {
      throw new Error(`exit ${code}`);
    },
  };

  let allowFrom = msteamsCfg.allowFrom;
  let groupAllowFrom = msteamsCfg.groupAllowFrom;
  let teamsConfig = msteamsCfg.teams;
  const allowNameMatching = isDangerousNameMatchingEnabled(msteamsCfg);

  const cleanAllowEntry = (entry: string) =>
    entry
      .replace(/^(msteams|teams):/i, "")
      .replace(/^user:/i, "")
      .trim();
  const isStableUserId = (entry: string) => /^[0-9a-fA-F-]{16,}$/.test(entry);
  const cleanAllowEntries = (entries?: string[]) =>
    entries?.map((entry) => cleanAllowEntry(entry)).filter((entry) => entry && entry !== "*") ?? [];
  const mergeStableUserIds = (entries?: string[]) => {
    const additions = cleanAllowEntries(entries).filter((entry) => isStableUserId(entry));
    return additions.length > 0 ? mergeAllowlist({ existing: entries, additions }) : entries;
  };

  const resolveAllowlistUsers = async (label: string, entries: string[]) => {
    if (entries.length === 0) {
      return { additions: [], unresolved: [] };
    }
    const resolved = await resolveMSTeamsUserAllowlist({ cfg, entries });
    const additions: string[] = [];
    const unresolved: string[] = [];
    for (const entry of resolved) {
      if (entry.resolved && entry.id) {
        additions.push(entry.id);
      } else {
        unresolved.push(entry.input);
      }
    }
    const mapping = resolved
      .filter((entry) => entry.resolved && entry.id)
      .map((entry) => `${entry.input}→${entry.id}`);
    summarizeMapping(label, mapping, unresolved, runtime);
    return { additions, unresolved };
  };

  try {
    allowFrom = mergeStableUserIds(allowFrom);
    if (Array.isArray(groupAllowFrom) && groupAllowFrom.length > 0) {
      groupAllowFrom = mergeStableUserIds(groupAllowFrom);
    }

    if (allowNameMatching) {
      const allowEntries = cleanAllowEntries(allowFrom).filter((entry) => !isStableUserId(entry));
      if (allowEntries.length > 0) {
        const { additions } = await resolveAllowlistUsers("msteams users", allowEntries);
        allowFrom = mergeAllowlist({ existing: allowFrom, additions });
      }

      if (Array.isArray(groupAllowFrom) && groupAllowFrom.length > 0) {
        const groupEntries = cleanAllowEntries(groupAllowFrom).filter(
          (entry) => !isStableUserId(entry),
        );
        if (groupEntries.length > 0) {
          const { additions } = await resolveAllowlistUsers("msteams group users", groupEntries);
          groupAllowFrom = mergeAllowlist({ existing: groupAllowFrom, additions });
        }
      }
    }

    if (teamsConfig && Object.keys(teamsConfig).length > 0) {
      const entries: Array<{ input: string; teamKey: string; channelKey?: string }> = [];
      for (const [teamKey, teamCfg] of Object.entries(teamsConfig)) {
        if (teamKey === "*") {
          continue;
        }
        const channels = teamCfg?.channels ?? {};
        const channelKeys = Object.keys(channels).filter((key) => key !== "*");
        if (channelKeys.length === 0) {
          entries.push({ input: teamKey, teamKey });
          continue;
        }
        for (const channelKey of channelKeys) {
          entries.push({
            input: `${teamKey}/${channelKey}`,
            teamKey,
            channelKey,
          });
        }
      }

      if (entries.length > 0) {
        const resolved = await resolveMSTeamsChannelAllowlist({
          cfg,
          entries: entries.map((entry) => entry.input),
        });
        const mapping: string[] = [];
        const unresolved: string[] = [];
        const nextTeams = { ...teamsConfig };

        resolved.forEach((entry, idx) => {
          const source = entries[idx];
          if (!source) {
            return;
          }
          const sourceTeam = teamsConfig?.[source.teamKey] ?? {};
          if (!entry.resolved || !entry.teamId) {
            unresolved.push(entry.input);
            return;
          }
          mapping.push(
            entry.channelId
              ? `${entry.input}→${entry.teamId}/${entry.channelId}`
              : `${entry.input}→${entry.teamId}`,
          );
          const existing = nextTeams[entry.teamId] ?? {};
          const mergedChannels = {
            ...sourceTeam.channels,
            ...existing.channels,
          };
          const mergedTeam = { ...sourceTeam, ...existing, channels: mergedChannels };
          nextTeams[entry.teamId] = mergedTeam;
          if (source.channelKey && entry.channelId) {
            const sourceChannel = sourceTeam.channels?.[source.channelKey];
            if (sourceChannel) {
              nextTeams[entry.teamId] = {
                ...mergedTeam,
                channels: {
                  ...mergedChannels,
                  [entry.channelId]: {
                    ...sourceChannel,
                    ...mergedChannels?.[entry.channelId],
                  },
                },
              };
            }
          }
        });

        teamsConfig = nextTeams;
        summarizeMapping("msteams channels", mapping, unresolved, runtime);
      }
    }
  } catch (err) {
    // Allowlist Graph resolution is security-sensitive — surface failures at
    // error level so operators notice the degraded state where Graph-resolved
    // IDs are missing (#77674).
    runtime.error?.(
      `msteams resolve failed; falling back to raw config entries — allowlist members resolved via Graph may be missing. ${formatUnknownError(err)}`,
    );
  }

  msteamsCfg = {
    ...msteamsCfg,
    allowFrom,
    groupAllowFrom,
    teams: teamsConfig,
  };
  cfg = {
    ...cfg,
    channels: {
      ...cfg.channels,
      msteams: msteamsCfg,
    },
  };

  const port = msteamsCfg.webhook?.port ?? 3978;
  const textLimit = core.channel.text.resolveTextChunkLimit(cfg, "msteams");
  const MB = 1024 * 1024;
  const agentDefaults = cfg.agents?.defaults;
  const mediaMaxBytes =
    typeof agentDefaults?.mediaMaxMb === "number" && agentDefaults.mediaMaxMb > 0
      ? Math.floor(agentDefaults.mediaMaxMb * MB)
      : 8 * MB;
  const conversationStore = opts.conversationStore ?? createMSTeamsConversationStoreFs();
  const pollStore = opts.pollStore ?? createMSTeamsPollStoreFs();

  log.info(`starting provider (port ${port})`);

  // Dynamic import to avoid loading SDK when provider is disabled
  const express = await import("express");

  // Create Express server first, then wrap it with the SDK's ExpressAdapter
  // so the App registers its route handler on it (including JWT validation).
  const expressApp = express.default();

  // Cap inbound webhook bodies before any handler parses them. The SDK's
  // ExpressAdapter installs `express.json()` (no limit) inside its registered
  // route, so an unbounded payload from a Bearer-shaped attacker would force
  // JSON parsing before JWT validation. Installing our own
  // `express.json({ limit })` first makes Express memoize the parsed body on
  // the request — the SDK's later `json()` is then a no-op, and our limit
  // applies before either the bearer gate or the SDK's authorize step run.
  expressApp.use(express.json({ limit: DEFAULT_WEBHOOK_MAX_BODY_BYTES }));

  // Cheap pre-parse auth gate: reject requests without a Bearer token before
  // spending CPU/memory on JSON body parsing. This prevents unauthenticated
  // request floods from forcing body parsing on internet-exposed webhooks.
  expressApp.use((req: Request, res: Response, next: (err?: unknown) => void) => {
    const auth = req.headers.authorization;
    if (!auth || !auth.startsWith("Bearer ")) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
    next();
  });

  const configuredPath = (msteamsCfg.webhook?.path ?? "/api/messages") as `/${string}`;

  // Lazy-load the SDK and create the App with ExpressAdapter. The SDK
  // registers POST /api/messages (or configured path) and handles JWT
  // validation + body parsing internally.
  const { app } = await loadMSTeamsSdkWithAuth(creds, {
    httpServerAdapter: await createMSTeamsExpressAdapter(expressApp),
    messagingEndpoint: configuredPath,
  });

  // Existing Azure Bot registrations may still point at the legacy
  // `/api/messages` endpoint while an operator has configured a custom
  // `webhook.path`. Forward to the configured path with a one-time deprecation
  // warning so those registrations keep working through the transition. The
  // forwarder runs after the SDK route is registered, so it only matches
  // requests that the SDK route itself didn't claim.
  if (configuredPath !== "/api/messages") {
    let warnedLegacyMessagesRoute = false;
    expressApp.post(
      "/api/messages",
      (req: Request, res: Response, next: (err?: unknown) => void) => {
        if (!warnedLegacyMessagesRoute) {
          warnedLegacyMessagesRoute = true;
          log.warn?.(
            `received request on /api/messages but webhook.path is ${configuredPath}; ` +
              "update your Azure Bot endpoint — this fallback will be removed in a future release",
          );
        }
        // Rewrite the URL so the SDK's registered handler picks it up. Express
        // app instances are themselves request handlers (Application extends
        // IRouter extends RequestHandler), so re-invoking the app re-runs the
        // middleware chain (including the SDK-registered route).
        req.url = configuredPath;
        expressApp(req, res, next);
      },
    );
  }

  // Build a token provider adapter for Graph API operations
  const tokenProvider = createMSTeamsTokenProvider(app);

  // Build SSO deps when the operator has opted in and a connection name
  // is configured. Leaving `sso` undefined matches the pre-SSO behavior
  // (the plugin will still ack signin invokes, but will not attempt a
  // Bot Framework token exchange or persist anything).
  let ssoDeps: MSTeamsSsoDeps | undefined;
  if (msteamsCfg.sso?.enabled && msteamsCfg.sso.connectionName) {
    ssoDeps = {
      tokenProvider,
      tokenStore: createMSTeamsSsoTokenStoreFs(),
      connectionName: msteamsCfg.sso.connectionName,
    };
    log.debug?.("msteams sso enabled", {
      connectionName: msteamsCfg.sso.connectionName,
    });
  }

  // Build a simple ActivityHandler-compatible object and register our
  // existing dispatch handlers on it. The SDK's App routes all inbound
  // activities to our handler via app.on('activity', ...).
  const handler = buildActivityHandler();
  const handlerDeps: MSTeamsMessageHandlerDeps = {
    cfg,
    runtime,
    appId,
    app,
    tokenProvider,
    textLimit,
    mediaMaxBytes,
    conversationStore,
    pollStore,
    log,
    sso: ssoDeps,
  };
  registerMSTeamsHandlers(handler, handlerDeps);

  // Handle adaptiveCard/action invokes (Action.Execute Universal Action Model).
  // We must return an InvokeResponse-shaped value so Teams updates the card UI;
  // returning nothing or letting the catch-all process it makes Teams report
  // "Unable to reach app".
  app.on("card.action", async (ctx): Promise<MSTeamsCardActionResponse> => {
    const adaptedCtx = adaptSdkContext(ctx, app);
    try {
      const activity = adaptedCtx.activity;
      const vote = extractMSTeamsPollVote(activity);
      if (vote) {
        const voterId = activity?.from?.aadObjectId ?? activity?.from?.id ?? "unknown";
        try {
          const poll = await pollStore.recordVote({
            pollId: vote.pollId,
            voterId,
            selections: vote.selections,
          });
          if (poll) {
            log.info("recorded poll vote", { pollId: vote.pollId, voterId });
            return {
              statusCode: 200,
              type: "application/vnd.microsoft.activity.message",
              value: "Vote recorded.",
            };
          }
          log.debug?.("poll vote ignored (poll not found)", { pollId: vote.pollId });
          return {
            statusCode: 200,
            type: "application/vnd.microsoft.activity.message",
            value: "Poll not found.",
          };
        } catch (err) {
          log.error("failed to record poll vote", {
            pollId: vote.pollId,
            error: formatUnknownError(err),
          });
          return {
            statusCode: 500,
            type: "application/vnd.microsoft.error",
            value: {
              code: "RECORD_VOTE_FAILED",
              message: "Could not record vote.",
              innerHttpError: { statusCode: 500, body: null },
            },
          };
        }
      }
      // Non-poll card actions: acknowledge silently and let the activity
      // catch-all dispatch the action data to the agent if applicable.
      await handler.run!(adaptedCtx);
      return {
        statusCode: 200,
        type: "application/vnd.microsoft.activity.message",
        value: "OK",
      };
    } catch (err) {
      log.error("msteams card.action failed", { error: formatUnknownError(err) });
      return {
        statusCode: 500,
        type: "application/vnd.microsoft.error",
        value: {
          code: "CARD_ACTION_FAILED",
          message: "Card action failed.",
          innerHttpError: { statusCode: 500, body: null },
        },
      };
    }
  });

  // File-consent invokes (large-file upload accept/decline). We register
  // typed handlers so the SDK writes the HTTP InvokeResponse for us — the
  // old `ctx.sendActivity({ type: "invokeResponse" })` shape no longer
  // works on the new SDK because that ctx call becomes an outbound BF
  // activity instead of the HTTP response (Brad #2 / codex #4).
  app.on("file.consent.accept", async (ctx) => {
    await runMSTeamsFileConsentInvokeHandler(adaptSdkContext(ctx, app), log);
  });
  app.on("file.consent.decline", async (ctx) => {
    await runMSTeamsFileConsentInvokeHandler(adaptSdkContext(ctx, app), log);
  });

  // SSO sign-in invokes. Registering these as user routes replaces the
  // SDK's built-in `signin.token-exchange` / `signin.verify-state` system
  // defaults — the SDK's router removes a default when a user route shares
  // its name. The SDK defaults would otherwise call
  // `api.users.token.exchange` themselves and emit a `signin` event that
  // nobody currently subscribes to. Our handler runs alone, persists the
  // token via `MSTeamsSsoTokenStore`, and the SDK wraps the void return
  // into the HTTP 200 InvokeResponse. The legacy
  // `ctx.sendActivity({ type: "invokeResponse", … })` ack is gone.
  app.on("signin.token-exchange", async (ctx) => {
    await runMSTeamsSigninInvokeHandler(adaptSdkContext(ctx, app), handlerDeps);
  });
  app.on("signin.verify-state", async (ctx) => {
    await runMSTeamsSigninInvokeHandler(adaptSdkContext(ctx, app), handlerDeps);
  });

  // Catch all inbound activities from the SDK and delegate to our existing
  // handler dispatch system. The SDK has already validated JWT and parsed the
  // activity by this point.
  app.on("activity", async (ctx) => {
    try {
      const adaptedCtx = adaptSdkContext(ctx, app);
      const activity = adaptedCtx.activity;
      // Skip invokes that have dedicated typed routes above.
      if (activity?.type === "invoke") {
        if (activity?.name === "adaptiveCard/action") {
          return;
        }
        if (activity?.name === "fileConsent/invoke") {
          return;
        }
        if (activity?.name === "signin/tokenExchange" || activity?.name === "signin/verifyState") {
          return;
        }
      }
      await handler.run!(adaptedCtx);
    } catch (err) {
      log.error("msteams webhook failed", { error: formatUnknownError(err) });
    }
  });

  // Initialize the SDK App — registers the POST route on Express and sets up
  // JWT validation middleware internally.
  await app.initialize();

  // Start listening and fail fast if bind/listen fails.
  const httpServer = expressApp.listen(port);
  await new Promise<void>((resolve, reject) => {
    const onListening = () => {
      httpServer.off("error", onError);
      log.info(`msteams provider started on port ${port}`);
      resolve();
    };
    const onError = (err: unknown) => {
      httpServer.off("listening", onListening);
      log.error("msteams server error", { error: formatUnknownError(err) });
      reject(err);
    };
    httpServer.once("listening", onListening);
    httpServer.once("error", onError);
  });
  applyMSTeamsWebhookTimeouts(httpServer);

  httpServer.on("error", (err) => {
    log.error("msteams server error", { error: formatUnknownError(err) });
  });

  const shutdown = async () => {
    log.info("shutting down msteams provider");
    return new Promise<void>((resolve) => {
      httpServer.close((err) => {
        if (err) {
          log.debug?.("msteams server close error", { error: formatUnknownError(err) });
        }
        resolve();
      });
    });
  };

  // Keep this task alive until close so gateway runtime does not treat startup as exit.
  await keepHttpServerTaskAlive({
    server: httpServer,
    abortSignal: opts.abortSignal,
    onAbort: shutdown,
  });

  return { app: expressApp, shutdown };
}

/**
 * Build a minimal ActivityHandler-compatible object that supports
 * onMessage / onMembersAdded registration and a run() method.
 */
function buildActivityHandler(): MSTeamsActivityHandler {
  type Handler = (context: unknown, next: () => Promise<void>) => Promise<void>;
  const messageHandlers: Handler[] = [];
  const membersAddedHandlers: Handler[] = [];
  const reactionsAddedHandlers: Handler[] = [];
  const reactionsRemovedHandlers: Handler[] = [];

  const handler: MSTeamsActivityHandler = {
    onMessage(cb) {
      messageHandlers.push(cb);
      return handler;
    },
    onMembersAdded(cb) {
      membersAddedHandlers.push(cb);
      return handler;
    },
    onReactionsAdded(cb) {
      reactionsAddedHandlers.push(cb);
      return handler;
    },
    onReactionsRemoved(cb) {
      reactionsRemovedHandlers.push(cb);
      return handler;
    },
    async run(context: unknown) {
      const ctx = context as { activity?: { type?: string } };
      const activityType = ctx?.activity?.type;
      const noop = async () => {};

      if (activityType === "message") {
        for (const h of messageHandlers) {
          await h(context, noop);
        }
      } else if (activityType === "conversationUpdate") {
        for (const h of membersAddedHandlers) {
          await h(context, noop);
        }
      } else if (activityType === "messageReaction") {
        const activity = (
          ctx as { activity?: { reactionsAdded?: unknown[]; reactionsRemoved?: unknown[] } }
        )?.activity;
        if (activity?.reactionsAdded?.length) {
          for (const h of reactionsAddedHandlers) {
            await h(context, noop);
          }
        }
        if (activity?.reactionsRemoved?.length) {
          for (const h of reactionsRemovedHandlers) {
            await h(context, noop);
          }
        }
      }
    },
  };

  return handler;
}

/**
 * Adapt a new @microsoft/teams.apps SDK context to the MSTeamsTurnContext interface
 * our handlers expect. The new SDK uses reply()/send() instead of sendActivity().
 */
function adaptSdkContext(ctx: unknown, app: MSTeamsApp): MSTeamsTurnContext {
  const sdkCtx = (ctx ?? {}) as {
    activity?: { id?: string; conversation?: { id?: string; conversationType?: string } };
    reply?: (activity: unknown) => Promise<unknown>;
    send?: (activity: unknown) => Promise<unknown>;
    stream?: {
      emit(a: unknown): void;
      update(t: string): void;
      close(): unknown;
      readonly canceled: boolean;
    };
  };
  if (typeof sdkCtx.reply !== "function" && typeof sdkCtx.send !== "function") {
    // Already adapted or old-style context — pass through.
    return ctx as MSTeamsTurnContext;
  }
  const conversationId = sdkCtx.activity?.conversation?.id ?? "";
  const conversationType = (sdkCtx.activity?.conversation?.conversationType ?? "").toLowerCase();
  const isThreadable = conversationType === "channel" || conversationType === "groupchat";
  // For Teams channels and group chats, use ctx.reply() so the SDK threads the
  // outbound activity to the inbound one (via replyToId + the inbound's
  // serviceUrl/conversation routing). For personal DMs, use ctx.send() instead
  // because reply() prepends a blockquote of the user's message — fine in
  // threaded surfaces where the visual nesting indicates context, but ugly in
  // 1:1 chat. Streaming chunks go through ctx.stream.emit/close separately.
  const sendActivity = (activity: unknown) =>
    isThreadable ? sdkCtx.reply!(activity) : sdkCtx.send!(activity);
  return Object.assign(Object.create(Object.getPrototypeOf(ctx)), ctx, {
    sendActivity,
    sendActivities: async (activities: unknown[]) => {
      const results: unknown[] = [];
      for (const a of activities) {
        results.push(await sendActivity(a));
      }
      return results;
    },
    updateActivity: async (activity: { id?: string; [key: string]: unknown }) => {
      const activityId = activity.id ?? "";
      return app.api.conversations.activities(conversationId).update(activityId, activity);
    },
    deleteActivity: async (activityId: string) => {
      return app.api.conversations.activities(conversationId).delete(activityId);
    },
    stream: sdkCtx.stream,
  });
}
