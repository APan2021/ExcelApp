# 使用官方的 OpenJDK 作为基础镜像
FROM openjdk:17-jdk-alpine

# 设置工作目录
WORKDIR /app

# 将当前目录下的jar文件复制到容器的/app目录
COPY target/excel-app-1.0-SNAPSHOT.jar /app/excel-app-1.0-SNAPSHOT.jar

# 暴露端口（根据你的Spring Boot应用配置的端口，一般为8080）
EXPOSE 8080

# 运行Spring Boot应用
ENTRYPOINT ["java", "-jar", "excel-app-1.0-SNAPSHOT.jar"]
