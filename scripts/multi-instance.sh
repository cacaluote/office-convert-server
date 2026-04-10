#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

APP_NAME="${APP_NAME:-office-convert-server}"
BINARY_PATH="${BINARY_PATH:-${SCRIPT_DIR}/${APP_NAME}}"
INSTANCE_COUNT="${INSTANCE_COUNT:-3}"
BASE_PORT="${BASE_PORT:-3000}"
HOST="${HOST:-0.0.0.0}"
RUN_ROOT="${RUN_ROOT:-${SCRIPT_DIR}/run}"
PID_DIR="${PID_DIR:-${RUN_ROOT}/pids}"
LOG_DIR="${LOG_DIR:-${RUN_ROOT}/logs}"
LIBREOFFICE_SDK_PATH_VALUE="${LIBREOFFICE_SDK_PATH:-}"
NO_AUTOMATIC_COLLECTION="${NO_AUTOMATIC_COLLECTION:-0}"
RUST_LOG_VALUE="${RUST_LOG:-info}"
COLOR_RED=$'\033[31m'
COLOR_GREEN=$'\033[32m'
COLOR_RESET=$'\033[0m'

usage() {
    cat <<EOF
Usage: $(basename "$0") <start|stop|restart|status>

Environment overrides:
  BINARY_PATH=<path>               Binary to launch
  INSTANCE_COUNT=<n>               Number of instances to manage (default: 3)
  BASE_PORT=<port>                 First port to bind (default: 3000)
  HOST=<host>                      Bind host (default: 0.0.0.0)
  RUN_ROOT=<dir>                   Root directory for logs and pid files
  PID_DIR=<dir>                    PID directory
  LOG_DIR=<dir>                    Log directory
  LIBREOFFICE_SDK_PATH=<path>      LibreOffice program directory
  NO_AUTOMATIC_COLLECTION=1        Disable post-request trim_memory calls
  RUST_LOG=<level/filter>          Log filter (default: info)

Examples:
  INSTANCE_COUNT=4 BASE_PORT=3100 ./scripts/multi-instance.sh start
  LIBREOFFICE_SDK_PATH=/usr/lib64/libreoffice/program ./scripts/multi-instance.sh restart
EOF
}

instance_port() {
    local index="$1"
    echo $((BASE_PORT + index))
}

instance_name() {
    local index="$1"
    printf '%s-%02d' "${APP_NAME}" "$((index + 1))"
}

pid_file() {
    local index="$1"
    echo "${PID_DIR}/$(instance_name "${index}").pid"
}

log_file() {
    local index="$1"
    echo "${LOG_DIR}/$(instance_name "${index}").log"
}

is_running() {
    local pid="$1"
    kill -0 "${pid}" 2>/dev/null
}

format_rss_kib() {
    local rss_kib="$1"
    local mib_whole=$((rss_kib / 1024))
    local mib_tenth=$((((rss_kib % 1024) * 10) / 1024))
    printf '%s.%s MiB RSS' "${mib_whole}" "${mib_tenth}"
}

memory_usage() {
    local pid="$1"
    local rss_kib

    rss_kib="$(ps -o rss= -p "${pid}" 2>/dev/null | tr -d '[:space:]')"
    if [[ -z "${rss_kib}" ]]; then
        echo "unknown RSS"
        return
    fi

    format_rss_kib "${rss_kib}"
}

ensure_dirs() {
    mkdir -p "${PID_DIR}" "${LOG_DIR}"
}

require_binary() {
    if [[ ! -x "${BINARY_PATH}" ]]; then
        echo "Binary is missing or not executable: ${BINARY_PATH}" >&2
        exit 1
    fi
}

build_command() {
    local port="$1"
    local -a cmd=("${BINARY_PATH}" "--host" "${HOST}" "--port" "${port}")

    if [[ -n "${LIBREOFFICE_SDK_PATH_VALUE}" ]]; then
        cmd+=("--office-path" "${LIBREOFFICE_SDK_PATH_VALUE}")
    fi

    if [[ "${NO_AUTOMATIC_COLLECTION}" == "1" ]]; then
        cmd+=("--no-automatic-collection" "true")
    fi

    printf '%q ' "${cmd[@]}"
}

start_instance() {
    local index="$1"
    local pid_path
    local log_path
    local port
    local name
    local cmd

    pid_path="$(pid_file "${index}")"
    log_path="$(log_file "${index}")"
    port="$(instance_port "${index}")"
    name="$(instance_name "${index}")"

    if [[ -f "${pid_path}" ]]; then
        local existing_pid
        existing_pid="$(<"${pid_path}")"
        if is_running "${existing_pid}"; then
            echo "${name} is already running on port ${port} (pid ${existing_pid})"
            return
        fi
        rm -f "${pid_path}"
    fi

    cmd="$(build_command "${port}")"
    echo "Starting ${name} on ${HOST}:${port}"
    nohup env RUST_LOG="${RUST_LOG_VALUE}" bash -lc "${cmd}" >>"${log_path}" 2>&1 &
    local pid=$!
    echo "${pid}" >"${pid_path}"
    echo "Started ${name} (pid ${pid}), log: ${log_path}"
}

stop_instance() {
    local index="$1"
    local pid_path
    local port
    local name

    pid_path="$(pid_file "${index}")"
    port="$(instance_port "${index}")"
    name="$(instance_name "${index}")"

    if [[ ! -f "${pid_path}" ]]; then
        echo "${name} is not running (no pid file)"
        return
    fi

    local pid
    pid="$(<"${pid_path}")"
    if is_running "${pid}"; then
        echo "Stopping ${name} on port ${port} (pid ${pid})"
        kill "${pid}"
        for _ in {1..20}; do
            if ! is_running "${pid}"; then
                break
            fi
            sleep 1
        done

        if is_running "${pid}"; then
            echo "${name} did not stop in time, sending SIGKILL"
            kill -9 "${pid}"
        fi
    else
        echo "${name} has stale pid file (${pid})"
    fi

    rm -f "${pid_path}"
}

status_instance() {
    local index="$1"
    local pid_path
    local port
    local name

    pid_path="$(pid_file "${index}")"
    port="$(instance_port "${index}")"
    name="$(instance_name "${index}")"

    if [[ ! -f "${pid_path}" ]]; then
        echo "${name}: ${COLOR_RED}stopped${COLOR_RESET} (port ${port})"
        return
    fi

    local pid
    pid="$(<"${pid_path}")"
    if is_running "${pid}"; then
        local memory
        memory="$(memory_usage "${pid}")"
        echo "${name}: ${COLOR_GREEN}running${COLOR_RESET} on ${HOST}:${port} (pid ${pid}, mem ${memory})"
    else
        echo "${name}: ${COLOR_RED}stale pid file${COLOR_RESET} on ${HOST}:${port} (pid ${pid})"
    fi
}

start_all() {
    ensure_dirs
    require_binary
    for ((i = 0; i < INSTANCE_COUNT; i++)); do
        start_instance "${i}"
    done
}

stop_all() {
    ensure_dirs
    for ((i = 0; i < INSTANCE_COUNT; i++)); do
        stop_instance "${i}"
    done
}

status_all() {
    ensure_dirs
    for ((i = 0; i < INSTANCE_COUNT; i++)); do
        status_instance "${i}"
    done
}

main() {
    local command="${1:-}"

    case "${command}" in
        start)
            start_all
            ;;
        stop)
            stop_all
            ;;
        restart)
            stop_all
            start_all
            ;;
        status)
            status_all
            ;;
        *)
            usage
            exit 1
            ;;
    esac
}

main "$@"
