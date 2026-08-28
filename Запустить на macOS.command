#!/bin/bash
set -u

fail() {
    printf '\n[ОШИБКА] %s\n' "$1" >&2
    printf 'Нажмите Enter, чтобы закрыть окно...'
    IFS= read -r _ || true
    exit 1
}

SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" 2>/dev/null && pwd -P)"
[[ -n "${SCRIPT_DIR:-}" ]] || fail "Не удалось определить папку приложения."
cd "$SCRIPT_DIR" || fail "Не удалось открыть папку приложения: $SCRIPT_DIR"

[[ -f "$SCRIPT_DIR/app.py" ]] || fail "Рядом с файлом запуска не найден app.py."
[[ -f "$SCRIPT_DIR/requirements.txt" ]] || fail "Рядом с файлом запуска не найден requirements.txt."

PYTHON_CMD=""
for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1 && \
       "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)' >/dev/null 2>&1; then
        PYTHON_CMD="$candidate"
        break
    fi
done

[[ -n "$PYTHON_CMD" ]] || \
    fail "Не найден Python 3.10 или новее. Установите его с https://www.python.org/downloads/."

VENV_DIR="$SCRIPT_DIR/.venv-macos"
VENV_PY="$VENV_DIR/bin/python"
if [[ ! -x "$VENV_PY" ]]; then
    printf 'Создаю виртуальное окружение для macOS...\n'
    "$PYTHON_CMD" -m venv "$VENV_DIR" || \
        fail "Не удалось создать виртуальное окружение."
fi

"$VENV_PY" -c 'import sys' >/dev/null 2>&1 || \
    fail "Виртуальное окружение повреждено. Удалите папку .venv-macos и запустите файл снова."

printf 'Проверяю и устанавливаю зависимости...\n'
"$VENV_PY" -m pip install --disable-pip-version-check -r "$SCRIPT_DIR/requirements.txt" || \
    fail "Не удалось установить зависимости. Проверьте интернет, прокси и совместимость версии Python."

printf '\nПриложение запускается. Не закрывайте это окно, пока оно работает.\n'
printf 'Если браузер не открылся, откройте адрес Local URL из этого окна.\n\n'
"$VENV_PY" -m streamlit run "$SCRIPT_DIR/app.py" --browser.gatherUsageStats=false
APP_EXIT_CODE=$?

if [[ $APP_EXIT_CODE -ne 0 && $APP_EXIT_CODE -ne 130 ]]; then
    fail "Streamlit завершился с кодом $APP_EXIT_CODE."
fi

exit 0
