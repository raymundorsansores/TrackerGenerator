web: java $JAVA_OPTS -jar target/dependency/webapp-runner.jar --port $PORT target/*.war
web: java -agentlib:jdwp=transport=dt_socket,server=y,address=9090,suspend=n -jar target/myapp.jar