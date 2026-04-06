FROM swift:6.2-jammy

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-core libreoffice-common ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY Package.swift Package.resolved ./
COPY Sources ./Sources
COPY Tests ./Tests
COPY Public ./Public

RUN swift build -c release

ENV PORT=8080
ENV STORAGE_ROOT=/data
ENV MAX_CONCURRENT_JOBS=3
ENV MAX_UPLOAD_SIZE_MB=80

EXPOSE 8080

CMD ["sh", "-lc", ".build/release/TourAutoLayoutWeb"]
