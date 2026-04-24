# Node.js + Ruby + Inkscape for MathType conversion
FROM node:20-bookworm-slim

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    ruby \
    ruby-dev \
    build-essential \
    curl \
    inkscape \
    ca-certificates \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Install Ruby gems
RUN gem install pry --no-document \
 && gem install mathtype_to_mathml --no-document \
 && (gem install mathtype_to_mathml_plus --no-document || true)

WORKDIR /app

# Copy package files and install
COPY package.json package-lock.json* ./
RUN npm install --production

# Copy application files
COPY server.js mt2mml.rb mt2mml_v2.rb ./

# Environment
ENV NODE_ENV=production
ENV PORT=8000

EXPOSE 8000

# Health check (server có /ping)
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
  CMD curl -f http://localhost:${PORT}/ping || exit 1

CMD ["node", "server.js"]
