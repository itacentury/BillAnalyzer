#!/bin/sh
set -e

# Ensure data directory has correct permissions
# This handles the case where a Docker volume is mounted as root
if [ -d "/data" ]; then
    chown -R appuser:appuser /data 2>/dev/null || true
fi

# Drop from root to appuser (with its supplementary groups) and run the command
exec setpriv --reuid appuser --regid appuser --init-groups "$@"
