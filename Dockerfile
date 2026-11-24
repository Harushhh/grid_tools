# The grid agent: Node 24 (the web app) + Python (the converters it runs).
# NOT BUILT OR RUN where this was written (no Docker there) - check it with `docker build -t grid-agent .` before relying on it.
#
#   docker run -p 3000:3000 -e GRID_APP_PASSWORD=choose-something-long -e COOKIE_SECURE=0 \
#              -v grid-data:/data -v "$PWD/master:/app/master" grid-agent
#
# COOKIE_SECURE=0 is only for trying it over plain http on your own computer; behind https leave it at 1 and set TRUST_PROXY=1 if a proxy sits in front.
FROM node:24-bookworm-slim

RUN apt-get update \
 && apt-get install -y --no-install-recommends python3 python3-venv \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# the converters' Python environment
COPY requirements.txt ./
RUN python3 -m venv .venv && .venv/bin/pip install --no-cache-dir -r requirements.txt

# the converters and the master lists (mount ./master over this to update the lists without a new image)
COPY *_to_config.py ./
COPY master ./master

# the web app
COPY webapp/package.json webapp/package-lock.json ./webapp/
RUN cd webapp && npm ci --omit=dev
COPY webapp ./webapp

RUN mkdir -p /data && chown -R node:node /data /app/webapp
ENV NODE_ENV=production \
    GRID_TOOLS_DIR=/app \
    DATA_DIR=/data \
    PORT=3000 \
    COOKIE_SECURE=1
VOLUME ["/data"]
EXPOSE 3000
USER node
WORKDIR /app/webapp
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s CMD node -e "fetch('http://127.0.0.1:'+(process.env.PORT||3000)+'/healthz').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"
CMD ["node", "--disable-warning=ExperimentalWarning", "server.js"]
