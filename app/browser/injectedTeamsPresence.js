/* Injected into the Teams page (main world). Mirrors presence calls from
 * fuk_u_teams.js: reads the MSAL secret from localStorage and PUTs
 * forceavailability. Idempotent — safe if did-finish-load runs more than once.
 *
 * Schedule JSON is inlined at inject time (see teamsPresenceInjection.js).
 */
(function () {
  const TAG = "[PRESENCE:page]";

  function pLog(...args) {
    console.info(TAG, ...args);
  }

  if (globalThis.__teamsForLinuxPresenceInjected) {
    pLog("already loaded, skipping duplicate inject");
    return;
  }
  globalThis.__teamsForLinuxPresenceInjected = true;

  pLog("script starting");

  /* eslint-disable-next-line no-undef -- inlined by teamsPresenceInjection.js */
  const SCHEDULE = __TEAMS_FOR_LINUX_PRESENCE_SCHEDULE__;

  const AVAILABLE = { availability: "Available" };
  const OFFLINE = { availability: "Offline", activity: "OffWork" };

  function parseHm(s) {
    const m = /^(\d{1,2}):(\d{2})$/.exec(s && String(s).trim());
    if (!m) {
      return null;
    }
    const h = Number(m[1]);
    const min = Number(m[2]);
    if (h > 23 || min > 59 || Number.isNaN(h) || Number.isNaN(min)) {
      return null;
    }
    return h * 60 + min;
  }

  const parsedStart = parseHm(SCHEDULE.workStart);
  const parsedEnd = parseHm(SCHEDULE.workEnd);
  const startMin = Number.isInteger(parsedStart) ? parsedStart : 9 * 60;
  const endMin = Number.isInteger(parsedEnd) ? parsedEnd : 19 * 60;
  const afterHoursMode = String(SCHEDULE.afterHours || "Offline")
    .trim()
    .toLowerCase();

  function minutesNowLocal() {
    const d = new Date();
    return d.getHours() * 60 + d.getMinutes();
  }

  function isWeekendLocal() {
    const dow = new Date().getDay();
    return dow === 0 || dow === 6;
  }

  function excludeWeekendsEnabled() {
    return SCHEDULE.excludeWeekends !== false;
  }

  function isWithinWorkHours() {
    if (excludeWeekendsEnabled() && isWeekendLocal()) {
      return false;
    }
    const now = minutesNowLocal();
    if (startMin < endMin) {
      return now >= startMin && now < endMin;
    }
    if (startMin > endMin) {
      return now >= startMin || now < endMin;
    }
    return false;
  }

  function getAuthToken() {
    for (const key in localStorage) {
      if (
        key.indexOf("https://presence.teams.microsoft.com//.default") >= 0
      ) {
        try {
          return JSON.parse(localStorage[key]).secret;
        } catch {
          return undefined;
        }
      }
    }
  }

  let presenceKeyLogged = false;
  function logLocalStoragePresenceHint() {
    if (presenceKeyLogged) {
      return;
    }
    presenceKeyLogged = true;
    let found = false;
    for (const key in localStorage) {
      if (key.indexOf("presence.teams.microsoft.com") >= 0) {
        found = true;
        pLog("localStorage key (truncated):", key.slice(0, 72) + "…");
      }
    }
    if (!found) {
      pLog(
        "no presence.teams.microsoft.com localStorage keys yet (login / token not ready)",
      );
    }
  }

  let httpLogCount = 0;

  function setStatus(status) {
    const token = getAuthToken();
    if (!token) {
      return;
    }
    const body = status ? JSON.stringify(status) : undefined;
    fetch("https://presence.teams.microsoft.com/v1/me/forceavailability/", {
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
      },
      body: body,
      method: "PUT",
    })
      .then(function (r) {
        httpLogCount += 1;
        if (httpLogCount <= 6 || !r.ok) {
          pLog("forceavailability HTTP", r.status, r.statusText);
        }
      })
      .catch(function (err) {
        console.error(TAG, "forceavailability fetch error:", err);
      });
  }

  let applyTickCount = 0;

  function applyScheduledPresence() {
    applyTickCount += 1;
    const token = getAuthToken();
    const nowM = minutesNowLocal();
    const inWork = isWithinWorkHours();
    if (!token) {
      logLocalStoragePresenceHint();
      if (applyTickCount === 1 || applyTickCount % 8 === 0) {
        pLog(
          "tick: no bearer token yet (tick #",
          applyTickCount,
          ") minutes=",
          nowM,
          "inWorkHours=",
          inWork,
        );
      }
      return;
    }
    if (inWork) {
      pLog("tick: work window -> Available; minutes=", nowM);
      setStatus(AVAILABLE);
      return;
    }
    const weekendOff =
      excludeWeekendsEnabled() && isWeekendLocal() && afterHoursMode === "offline";
    if (afterHoursMode === "offline") {
      pLog(
        weekendOff
          ? "tick: weekend -> Offline; minutes="
          : "tick: outside work window -> Offline; minutes=",
        nowM,
      );
      setStatus(OFFLINE);
    } else {
      pLog(
        excludeWeekendsEnabled() && isWeekendLocal()
          ? "tick: weekend; afterHours=None, not calling API; minutes="
          : "tick: outside work window; afterHours=None, not calling API; minutes=",
        nowM,
      );
    }
  }

  pLog("schedule", SCHEDULE, {
    startMin,
    endMin,
    afterHoursMode,
    excludeWeekends: excludeWeekendsEnabled(),
    weekendToday: isWeekendLocal(),
    nowMinutes: minutesNowLocal(),
    inWorkHours: isWithinWorkHours(),
  });

  applyScheduledPresence();

  let bootAttempts = 0;
  const bootInterval = setInterval(function () {
    bootAttempts += 1;
    if (getAuthToken()) {
      pLog("boot: token appeared after", bootAttempts, "attempts (~2s each)");
      applyScheduledPresence();
      clearInterval(bootInterval);
    } else if (bootAttempts >= 90) {
      pLog("boot: giving up after 90 attempts (~3 min) without presence token");
      clearInterval(bootInterval);
    } else if (bootAttempts % 15 === 0) {
      pLog("boot: still waiting for presence token, attempt", bootAttempts);
    }
  }, 2000);

  setInterval(applyScheduledPresence, 15 * 1000);
  pLog("intervals armed: boot poll 2s, apply 15s");
})();
