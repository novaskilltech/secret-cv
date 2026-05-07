#!/bin/bash
set -e

# Cleanup any orphaned files from previous runs
echo "Cleaning up /tmp/nova..."
rm -rf /tmp/nova/* || true
mkdir -p /tmp/nova
chmod 1777 /tmp/nova

# Start the application
echo "Starting NOVA PDF Backend..."
exec "$@"
