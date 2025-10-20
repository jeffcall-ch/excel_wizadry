#!/usr/bin/env bash
set -euo pipefail

# === Config ===
CONCURRENCY=3
DEST_REMOTE="onedrive:Domi_videok"
WORKDIR="${WORKDIR:-$PWD}"         # default: current folder
COOKIES="${COOKIES:-$WORKDIR/cookies.txt}"
URLS_FILE="${URLS_FILE:-$WORKDIR/videos.txt}"
LOGFILE="${LOGFILE:-$WORKDIR/download_to_onedrive.log}"

# Output template: safe filenames for OneDrive (strip illegal chars)
OUT_TPL="%(title,replace='[\\/:*?\"<>|]','_')s.%(ext)s"

echo "== $(date '+%F %T') | Start ==" | tee -a "$LOGFILE"
echo "WORKDIR: $WORKDIR" | tee -a "$LOGFILE"
echo "DEST_REMOTE: $DEST_REMOTE" | tee -a "$LOGFILE"

need() { command -v "$1" >/dev/null 2>&1 || { echo "Missing dependency: $1" | tee -a "$LOGFILE"; exit 1; }; }
need yt-dlp
need rclone
need stdbuf || true   # usually in coreutils; if absent, output still works, just less line-buffered

[[ -f "$COOKIES" ]] || { echo "cookies.txt not found at: $COOKIES" | tee -a "$LOGFILE"; exit 1; }
[[ -f "$URLS_FILE" ]] || { echo "videos.txt not found at: $URLS_FILE" | tee -a "$LOGFILE"; exit 1; }

# Ensure destination exists
if ! rclone lsd "$DEST_REMOTE" >/dev/null 2>&1; then
  echo "Creating remote folder: $DEST_REMOTE" | tee -a "$LOGFILE"
  rclone mkdir "$DEST_REMOTE"
fi

# Load URLs (strip blanks & comments)
mapfile -t URLS < <(sed -e 's/^[[:space:]]*//; s/[[:space:]]*$//' -e '/^$/d' -e '/^#/d' "$URLS_FILE")
COUNT=${#URLS[@]}
if (( COUNT == 0 )); then
  echo "No URLs found in $URLS_FILE" | tee -a "$LOGFILE"
  exit 0
fi
echo "Found $COUNT URL(s) in $URLS_FILE" | tee -a "$LOGFILE"

# Debug: show what we loaded
echo "Debug: Loaded URLs:" | tee -a "$LOGFILE"
for i in "${!URLS[@]}"; do
  echo "  [$i] ${URLS[$i]}" | tee -a "$LOGFILE"
done

# Download one URL and move to OneDrive, with tagged progress
download_one() {
  local job_id="$1"
  local url="$2"
  local tag="[job ${job_id}]"

  echo "$tag START $(date '+%T') -> $url" | tee -a "$LOGFILE"

  # --progress + --newline produce continuous progress lines
  # stdbuf forces line-buffered output to keep lines intact across jobs
  # Filter progress to show only: START/DONE, errors, warnings, and every 10% milestone
  if command -v stdbuf >/dev/null 2>&1; then
    stdbuf -oL -eL yt-dlp --cookies "$COOKIES" \
           --no-part \
           --restrict-filenames \
           --concurrent-fragments 1 \
           --paths home:"$WORKDIR" \
           -o "$OUT_TPL" \
           --sleep-requests 1 \
           --min-sleep-interval 1 --max-sleep-interval 5 \
           --progress --newline \
           --retries infinite --fragment-retries 10 \
           --exec "rclone move -v '{}'* '$DEST_REMOTE' --transfers 4 --checkers 8" \
           "$url" 2>&1 \
      | awk -v tag="$tag" '
          /\[download\]/ && /[0-9]+\.[0-9]+%/ {
            # Extract percentage
            match($0, /[0-9]+\.[0-9]+%/, pct)
            p = pct[0]
            sub(/%/, "", p)
            # Show every 10% and 100%
            if (int(p) % 10 == 0 && int(p) != last_pct) {
              print tag, $0
              fflush()
              last_pct = int(p)
            }
            next
          }
          # Show all non-progress lines (errors, warnings, etc.)
          { print tag, $0; fflush() }
        ' \
      | tee -a "$LOGFILE"
  else
    yt-dlp --cookies "$COOKIES" \
           --no-part \
           --restrict-filenames \
           --concurrent-fragments 1 \
           --paths home:"$WORKDIR" \
           -o "$OUT_TPL" \
           --sleep-requests 1 \
           --min-sleep-interval 1 --max-sleep-interval 5 \
           --progress --newline \
           --retries infinite --fragment-retries 10 \
           --exec "rclone move -v '{}'* '$DEST_REMOTE' --transfers 4 --checkers 8" \
           "$url" 2>&1 \
      | awk -v tag="$tag" '
          /\[download\]/ && /[0-9]+\.[0-9]+%/ {
            # Extract percentage
            match($0, /[0-9]+\.[0-9]+%/, pct)
            p = pct[0]
            sub(/%/, "", p)
            # Show every 10% and 100%
            if (int(p) % 10 == 0 && int(p) != last_pct) {
              print tag, $0
              fflush()
              last_pct = int(p)
            }
            next
          }
          # Show all non-progress lines (errors, warnings, etc.)
          { print tag, $0; fflush() }
        ' \
      | tee -a "$LOGFILE"
  fi

  echo "$tag DONE  $(date '+%T') -> $url" | tee -a "$LOGFILE"
}

# Concurrency control (no GNU parallel needed)
echo "Debug: Entering download loop" | tee -a "$LOGFILE"
active_jobs=0
job_id=0
for url in "${URLS[@]}"; do
  job_id=$((job_id + 1))
  echo "Debug: Processing URL #$job_id" | tee -a "$LOGFILE"

  # Wait until a slot is free
  while (( active_jobs >= CONCURRENCY )); do
    wait -n || true
    active_jobs=$(jobs -r | wc -l)
  done

  download_one "$job_id" "$url" &
  active_jobs=$((active_jobs + 1))
  echo "Debug: Launched job $job_id, active: $active_jobs" | tee -a "$LOGFILE"
done
echo "Debug: Loop finished, waiting for jobs" | tee -a "$LOGFILE"

wait

echo "== $(date '+%F %T') | All done. Logs: $LOGFILE ==" | tee -a "$LOGFILE"
