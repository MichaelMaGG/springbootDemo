FROM openjdk:8-jdk-alpine
MAINTAINER michaelggmanz@gmail.com
VOLUME /tmp
ADD target/*.jar app.jar
ENTRYPOINT ["java", "-jar", "/app.jar"]
