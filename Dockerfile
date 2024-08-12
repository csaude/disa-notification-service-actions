FROM eclipse-temurin:11-jre-alpine
RUN ln -s /usr/share/zoneinfo/Africa/Maputo /etc/localtime
RUN mkdir /opt/disa-notification-service
COPY target/notification.service-2.6.1-SNAPSHOT.jar /opt/disa-notification-service/
CMD ["java", "-jar", "/opt/disa-notification-service/notification.service-2.6.1-SNAPSHOT.jar"]
