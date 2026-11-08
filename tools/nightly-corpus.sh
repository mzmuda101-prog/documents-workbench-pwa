#!/bin/zsh
# nightly-corpus.sh — nocny przebieg niezmienników na korpusie (scripts/corpus-invariants-playwright.js).
#
# Uruchamiany przez launchd (tools/com.dwb.nightly-corpus.plist, codziennie 3:00; przegapiony —
# po wybudzeniu Maca). Ziarno losowych edycji = data, więc co noc inne ścieżki, a raport podaje
# polecenie do powtórzenia. Raporty: output/corpus/<data>-<silnik>.md, log: output/corpus/nightly.log.
#
#   ręcznie:   zsh tools/nightly-corpus.sh
#   wyłączyć:  launchctl bootout gui/$(id -u)/com.dwb.nightly-corpus

ROOT="${0:A:h:h}"
cd "$ROOT" || exit 1
export PATH="$HOME/.nvm/versions/node/v26.13.0/bin:/opt/homebrew/bin:/usr/bin:/bin:$PATH"
mkdir -p output/corpus
LOG=output/corpus/nightly.log
SEED=$(date +%Y%m%d)
echo "=== $(date '+%Y-%m-%d %H:%M') SEED=$SEED" >> "$LOG"

# serwer testowy (ten sam co dla testów) — gdy nie działa, uruchamiamy na czas przebiegu
SERVER_PID=""
if ! curl -s -o /dev/null http://127.0.0.1:7823/; then
  python3 scripts/test-server.py 7823 >/dev/null 2>&1 &
  SERVER_PID=$!
  sleep 2
fi

STATUS=0
for ENGINE in chromium webkit; do
  ENGINE=$ENGINE SEED=$SEED STEPS=40 node scripts/corpus-invariants-playwright.js >> "$LOG" 2>&1 || STATUS=1
done
[ -n "$SERVER_PID" ] && kill "$SERVER_PID" 2>/dev/null

if [ $STATUS -ne 0 ]; then
  osascript -e "display notification \"Niezmienniki: są błędy — output/corpus/$(date +%Y-%m-%d)-*.md\" with title \"Documents Workbench\"" 2>/dev/null
fi
echo "=== koniec, status $STATUS" >> "$LOG"
exit $STATUS
