# Keep the renderer and its fonts identical on developer machines and CI.
FROM ubuntu:24.04@sha256:33ceb71981b602c1a7443a53469e4dba065f7503eab3078a2d7a57a2ab987517

# https://snapshot.ubuntu.com/ documents the immutable package archive.
RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends ca-certificates \
    && apt-get update --snapshot 20260903T000000Z -o APT::Update::Error-Mode=any \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      --snapshot 20260903T000000Z \
      fonts-liberation libreoffice-writer poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Math must be installed explicitly when recommended packages are disabled.
RUN apt-get update --snapshot 20260903T000000Z -o APT::Update::Error-Mode=any \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      --snapshot 20260903T000000Z \
      fonts-crosextra-carlito fonts-dejavu-core libreoffice-math \
    && rm -rf /var/lib/apt/lists/*

COPY fonts.conf /etc/fonts/conf.d/99-docx-kit.conf
