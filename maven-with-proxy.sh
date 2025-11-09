#!/bin/bash

# Extract proxy details from environment variable
PROXY_URL="$https_proxy"

# Parse proxy URL using parameter expansion
# Format: http://username:password@host:port
PROXY_WITHOUT_PROTOCOL="${PROXY_URL#http://}"
PROXY_AUTH="${PROXY_WITHOUT_PROTOCOL%@*}"
PROXY_HOST_PORT="${PROXY_WITHOUT_PROTOCOL##*@}"
PROXY_USER="${PROXY_AUTH%:*}"
PROXY_PASS="${PROXY_AUTH#*:}"
PROXY_HOST="${PROXY_HOST_PORT%:*}"
PROXY_PORT="${PROXY_HOST_PORT##*:}"

echo "Proxy Host: $PROXY_HOST"
echo "Proxy Port: $PROXY_PORT"

# Run Maven with proxy settings
MAVEN_OPTS="-Dhttp.proxyHost=$PROXY_HOST -Dhttp.proxyPort=$PROXY_PORT -Dhttp.proxyUser=$PROXY_USER -Dhttp.proxyPassword=$PROXY_PASS -Dhttps.proxyHost=$PROXY_HOST -Dhttps.proxyPort=$PROXY_PORT -Dhttps.proxyUser=$PROXY_USER -Dhttps.proxyPassword=$PROXY_PASS -Dhttp.nonProxyHosts=localhost|127.0.0.1|169.254.169.254" mvn "$@"
