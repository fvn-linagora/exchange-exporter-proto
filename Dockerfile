FROM linagofab/fsprojectscaffold

MAINTAINER Fabien <fvignon@linagora.com>

RUN mkdir -p /config

VOLUME ["/config"]

CMD ["/binaries/EchangeExporterProto/EchangeExporterProto.exe", "--help"]
