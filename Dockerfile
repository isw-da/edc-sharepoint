FROM eclipse-temurin:17-jre

# Run as a non-root numeric UID. We don't create a user record — Java
# doesn't need one, and avoiding useradd sidesteps base-image UID/GID
# collisions. The application opens no Unix sockets or files that need
# named ownership.
COPY target/connector-server-sharepoint-1.0.0-exec.jar /opt/connector-server-sharepoint.jar

USER 1000:1000

EXPOSE 7339

# No Dockerfile HEALTHCHECK — Kubernetes liveness/readiness probes are
# the right mechanism in this deployment model. See the pod manifest in
# README.md.

CMD ["java", "-Duser.timezone=UTC", "-jar", "/opt/connector-server-sharepoint.jar"]
