name := "Kafka-practice"

version := "0.1"

scalaVersion := "2.12.0"

libraryDependencies ++= List(
  "org.apache.kafka" %% "kafka" % "2.1.0",
  "com.typesafe.akka" %% "akka-actor" % "2.5.24",
  "org.apache.poi" % "poi-ooxml" % "3.17"

)