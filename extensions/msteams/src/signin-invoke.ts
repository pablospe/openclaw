import { formatUnknownError } from "./errors.js";
import { isSigninInvokeAuthorized } from "./monitor-handler.js";
import type { MSTeamsMessageHandlerDeps } from "./monitor-handler.types.js";
import type { MSTeamsTurnContext } from "./sdk-types.js";
import {
  handleSigninTokenExchangeInvoke,
  handleSigninVerifyStateInvoke,
  parseSigninTokenExchangeValue,
  parseSigninVerifyStateValue,
} from "./sso.js";

/**
 * Run the SSO sign-in invoke handler for `signin/tokenExchange` and
 * `signin/verifyState`.
 *
 * Our typed `app.on("signin.token-exchange" | "signin.verify-state", …)`
 * registrations replace the SDK's built-in system defaults — the SDK's
 * router removes the default when a user route shares its name. Only our
 * handler runs; the SDK wraps a `void` return into `{ status: 200 }`. The
 * legacy `ctx.sendActivity({ type: "invokeResponse", … })` ack is gone —
 * on the new SDK that becomes an outbound BF activity instead of the HTTP
 * response.
 */
export async function runMSTeamsSigninInvokeHandler(
  context: MSTeamsTurnContext,
  deps: MSTeamsMessageHandlerDeps,
): Promise<void> {
  const activity = context.activity;

  if (!(await isSigninInvokeAuthorized(context, deps))) {
    return;
  }

  if (!deps.sso) {
    deps.log.debug?.("signin invoke received but msteams.sso is not configured", {
      name: activity.name,
    });
    return;
  }

  const user = {
    userId: activity.from?.aadObjectId ?? activity.from?.id ?? "",
    channelId: activity.channelId ?? "msteams",
  };

  try {
    if (activity.name === "signin/tokenExchange") {
      const parsed = parseSigninTokenExchangeValue(activity.value);
      if (!parsed) {
        deps.log.debug?.("invalid signin/tokenExchange invoke value");
        return;
      }
      const result = await handleSigninTokenExchangeInvoke({
        value: parsed,
        user,
        deps: deps.sso,
      });
      if (result.ok) {
        deps.log.info("msteams sso token exchanged", {
          userId: user.userId,
          hasExpiry: Boolean(result.expiresAt),
        });
      } else {
        deps.log.error("msteams sso token exchange failed", {
          code: result.code,
          status: result.status,
          message: result.message,
        });
      }
      return;
    }

    // signin/verifyState
    const parsed = parseSigninVerifyStateValue(activity.value);
    if (!parsed) {
      deps.log.debug?.("invalid signin/verifyState invoke value");
      return;
    }
    const result = await handleSigninVerifyStateInvoke({
      value: parsed,
      user,
      deps: deps.sso,
    });
    if (result.ok) {
      deps.log.info("msteams sso verifyState succeeded", {
        userId: user.userId,
        hasExpiry: Boolean(result.expiresAt),
      });
    } else {
      deps.log.error("msteams sso verifyState failed", {
        code: result.code,
        status: result.status,
        message: result.message,
      });
    }
  } catch (err) {
    deps.log.error("msteams sso invoke handler error", {
      error: formatUnknownError(err),
    });
  }
}
