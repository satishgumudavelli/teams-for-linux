const fs = require("node:fs");
const path = require("node:path");
const log = require("electron-log/main");

const SCRIPT_PATH = path.join(__dirname, "injectedTeamsPresence.js");

const PLACEHOLDER = "__TEAMS_FOR_LINUX_PRESENCE_SCHEDULE__";

/** @type {WeakSet<Electron.WebContents>} */
const presenceConsoleMirrored = new WeakSet();

/**
 * Forwards injected page logs to electron-log so they appear in the terminal.
 * @param {Electron.WebContents} webContents
 */
function attachPresenceConsoleMirror(webContents) {
  if (!webContents || presenceConsoleMirrored.has(webContents)) {
    return;
  }
  presenceConsoleMirrored.add(webContents);
  webContents.on("console-message", (event) => {
    const message = typeof event?.message === "string" ? event.message : "";
    if (!message.includes("[PRESENCE:page]")) {
      return;
    }
    log.info("[PRESENCE:renderer]", message);
  });
}

/**
 * Runs the Teams presence injection in the given webContents when enabled in config.
 * @param {Electron.WebContents} webContents
 * @param {{
 *   forceAvailablePresenceOnLoad?: boolean,
 *   forceAvailablePresenceTimeStart?: string,
 *   forceAvailablePresenceTimeEnd?: string,
 *   forceAvailablePresenceAfterHours?: string,
 *   forceAvailablePresenceExcludeWeekends?: boolean,
 * }} config
 */
function injectTeamsPresence(webContents, config) {
  if (!config?.forceAvailablePresenceOnLoad) {
    log.debug("[PRESENCE] skip: forceAvailablePresenceOnLoad is false");
    return;
  }
  if (!webContents || webContents.isDestroyed()) {
    log.warn("[PRESENCE] skip: webContents missing or destroyed");
    return;
  }

  attachPresenceConsoleMirror(webContents);

  try {
    let source = fs.readFileSync(SCRIPT_PATH, "utf8");
    const schedule = {
      workStart: config.forceAvailablePresenceTimeStart ?? "09:00",
      workEnd: config.forceAvailablePresenceTimeEnd ?? "19:00",
      afterHours: config.forceAvailablePresenceAfterHours ?? "Offline",
      excludeWeekends: config.forceAvailablePresenceExcludeWeekends !== false,
    };
    const embedded = JSON.stringify(schedule);
    if (!source.includes(PLACEHOLDER)) {
      log.error(
        "[PRESENCE] injectedTeamsPresence.js missing schedule placeholder",
      );
      return;
    }
    source = source.replaceAll(PLACEHOLDER, embedded);
    if (source.includes(PLACEHOLDER)) {
      log.error("[PRESENCE] placeholder still present after replaceAll");
      return;
    }
    let pageUrl = "";
    try {
      pageUrl = webContents.getURL();
    } catch {
      pageUrl = "(getURL failed)";
    }
    log.info("[PRESENCE] injecting page script", { schedule, pageUrl });
    webContents
      .executeJavaScript(source, true)
      .then(() => {
        log.info(
          "[PRESENCE] executeJavaScript completed (page script ran without throw)",
        );
      })
      .catch((err) => {
        log.error("[PRESENCE] executeJavaScript failed:", err.message);
      });
  } catch (err) {
    log.error("[PRESENCE] read/prepare injectedTeamsPresence.js failed:", err);
  }
}

module.exports = { injectTeamsPresence };
