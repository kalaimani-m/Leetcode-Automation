"""
Thin client for LeetCode's public (unofficial) GraphQL endpoint.

LeetCode has no official public API. This uses the same endpoint LeetCode's
own website calls (leetcode.com/graphql), requesting only public profile data.
No login/API key is needed for this, but it is undocumented and can change
or rate-limit without notice -- hence the retries/backoff and the N/A
fallbacks used throughout main.py.

IMPORTANT LIMITATION (read this before trusting "today's" numbers blindly):
LeetCode's recentAcSubmissionList only returns a student's most recent
~15-20 ACCEPTED submissions, not a full log. If a student accepts more than
that many problems today, or has other accepted submissions between their
last one today and "now" pushing older ones out of the window, this will
UNDERCOUNT today's unique solved problems. There is no public endpoint that
returns a clean "accepted problems on date X" list. This script flags such
cases in the Notes sheet of the report rather than silently guessing.
"""

import time
import re
import requests
from datetime import datetime
from zoneinfo import ZoneInfo

GRAPHQL_URL = "https://leetcode.com/graphql"
IST = ZoneInfo("Asia/Kolkata")

HEADERS = {
    "Content-Type": "application/json",
    "Referer": "https://leetcode.com",
    "User-Agent": "Mozilla/5.0 (compatible; StudentProgressBot/1.0)",
}

QUERY = """
query getUserProfile($username: String!) {
  matchedUser(username: $username) {
    username
    submitStatsGlobal {
      acSubmissionNum {
        difficulty
        count
      }
    }
  }
  recentAcSubmissionList(username: $username, limit: 20) {
    title
    titleSlug
    timestamp
  }
}
"""


def extract_username(profile_url):
    """Pull the username out of a leetcode.com/u/<username>/ URL."""
    if not profile_url or not isinstance(profile_url, str):
        return None
    match = re.search(r"leetcode\.com/u/([^/]+)/?", profile_url.strip())
    if match:
        return match.group(1)
    # fallback for the older /<username>/ profile URL format
    match = re.search(r"leetcode\.com/([^/]+)/?$", profile_url.strip())
    return match.group(1) if match else None


def fetch_profile(username, max_retries=3, base_delay=3):
    """
    Returns a dict:
      {
        "overall_solved": int or None,
        "recent_ac": [ {title, titleSlug, timestamp}, ... ] or None,
        "status": "ok" | "not_found" | "error",
        "error": str or None,
      }
    """
    if not username:
        return {"overall_solved": None, "recent_ac": None,
                "status": "not_found", "error": "No username in profile URL"}

    payload = {"query": QUERY, "variables": {"username": username}}
    last_error = None

    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.post(GRAPHQL_URL, json=payload, headers=HEADERS, timeout=15)

            if resp.status_code == 429:
                wait = base_delay * (2 ** (attempt - 1))
                time.sleep(wait)
                continue

            resp.raise_for_status()
            body = resp.json()

            if body.get("errors"):
                last_error = str(body["errors"])
                time.sleep(base_delay)
                continue

            data = body.get("data") or {}
            matched = data.get("matchedUser")

            if matched is None:
                return {"overall_solved": None, "recent_ac": None,
                        "status": "not_found", "error": "Profile not found or private"}

            overall = None
            for entry in matched.get("submitStatsGlobal", {}).get("acSubmissionNum", []):
                if entry.get("difficulty") == "All":
                    overall = entry.get("count")
                    break

            recent_ac = data.get("recentAcSubmissionList")

            return {"overall_solved": overall, "recent_ac": recent_ac,
                    "status": "ok", "error": None}

        except requests.exceptions.RequestException as exc:
            last_error = str(exc)
            time.sleep(base_delay * attempt)

    return {"overall_solved": None, "recent_ac": None,
            "status": "error", "error": last_error or "Failed after retries"}


def count_unique_solved_today(recent_ac, today_ist_date=None):
    """
    Dedupe accepted submissions by titleSlug and count only those whose
    timestamp falls on today's IST calendar date.

    Returns (count, capped) where capped=True means the recent list was
    full (>=20 entries) and the true count for today may be higher than
    what's reported -- caller should flag this in the report.
    """
    if recent_ac is None:
        return None, False

    if today_ist_date is None:
        today_ist_date = datetime.now(IST).date()

    seen_slugs = set()
    for sub in recent_ac:
        ts = int(sub["timestamp"])
        sub_date = datetime.fromtimestamp(ts, IST).date()
        if sub_date == today_ist_date:
            seen_slugs.add(sub["titleSlug"])

    capped = len(recent_ac) >= 20
    return len(seen_slugs), capped
