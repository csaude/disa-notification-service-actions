## DISA notification service

Spring boot app for periodic sending of DISA EPTS interoperability reports.

## Pre-requisites
- JDK 11
- Running instance of [disa-api](https://github.com/FriendsInGlobalHealth/disa-api).

## Building
Run `mvn clean install`

## Running
1.  Provide environment variables required by the `src/main/resources/application.properties` file.

    This can be done by creating a `.env` file:

    ```bash
    export DISA_DATASOURCE_URL=
    export DISA_DATASOURCE_USERNAME=
    export DISA_DATASOURCE_PASSWORD=
    export DISA_MAIL_PROTOCOL=
    export DISA_MAIL_HOST=
    export DISA_MAIL_PORT=
    export DISA_MAIL_USERNAME=
    export DISA_MAIL_PASSWORD=
    ```

2.  Source the `.env` file and execute the built `.jar`:

    ```bash
    source 'filename.env' && java -jar notification-service.jar
    ```

## Generating reports on disk
FileSystemMailService is an implementation of MailService that allows generating the reports without
sending to recipients.
It is configured by setting the property `app.mailservice` to `fileSystem`, either on
application.properties or by supplying a param to the excutable jar:

   ```bash
    java -jar notification-service.jar --app.mailservice=fileSystem
    ```
