#!/usr/bin/env bash
# Проверка: общие reference-файлы во всех скиллах идентичны (нет дрейфа копий).
# Использование: scripts/check_shared_refs.sh        — только проверка
#                scripts/check_shared_refs.sh --sync — скопировать эталон во все копии
set -euo pipefail
cd "$(dirname "$0")/../skills"

# файл : эталонный скилл (источник правды)
declare -A MASTER=(
  [budget-limits.md]=acivs-knowledge
  [signatories.md]=dkt-letters-internal
  [kd-clause-map.md]=acivs-knowledge
)
# скиллы, где обязана лежать копия (если файла нет — создаётся при --sync)
REQUIRED_budget_limits=(acivs-knowledge acivs-budget-check acivs-red-flags dogovor-navigator dogovor-smeta-check)
REQUIRED_kd_clause_map=("${REQUIRED_budget_limits[@]}")

rc=0
for f in "${!MASTER[@]}"; do
  src="${MASTER[$f]}/references/$f"
  [[ -f "$src" ]] || { echo "НЕТ эталона: $src"; rc=1; continue; }
  req="REQUIRED_${f%.md}"; req="${req//-/_}"
  declare -n reqs="$req" 2>/dev/null || reqs=()
  for sk in "${reqs[@]}"; do
    [[ -f "$sk/references/$f" ]] && continue
    if [[ "${1:-}" == "--sync" ]]; then cp "$src" "$sk/references/$f"; echo "создано: $sk/references/$f"
    else echo "ОТСУТСТВУЕТ: $sk/references/$f"; rc=1; fi
  done
  unset -n reqs 2>/dev/null || true
  for copy in */references/"$f"; do
    [[ "$copy" == "$src" ]] && continue
    if ! cmp -s "$src" "$copy"; then
      if [[ "${1:-}" == "--sync" ]]; then cp "$src" "$copy"; echo "синхронизировано: $copy"
      else echo "РАСХОЖДЕНИЕ: $copy != $src"; rc=1; fi
    fi
  done
done
[[ $rc -eq 0 ]] && echo "OK: общие справочники синхронны"
exit $rc
