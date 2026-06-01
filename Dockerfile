FROM eclipse-temurin:17-jre
COPY target/connector-server-sharepoint-1.0.0-exec.jar /opt/connector-server-sharepoint.jar
EXPOSE 7339
CMD ["java", "-Duser.timezone=UTC", "-jar", "/opt/connector-server-sharepoint.jar"]
