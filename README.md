# Read/Write Excel files with Spring and Apache POI

Working with **Excel** documents is a frequently used feature in a software application.
[Apache POI](https://poi.apache.org/) is a popular open source library run by the Apache Software Foundation which is developed for reading and writing files in Microsoft Office formats, such as Word, PowerPoint, and Excel.
This project demonstrates the use of the **Apache POI** for working with **Excel** spreadsheets.

## Overview

The Apache POI library supports both `.xls` and `.xlsx` files and is a more complex library than other Java libraries for working with Excel files.
It provides the *Workbook* interface for modeling an *Excel* file, and the *Sheet*, *Row* and *Cell* interfaces that model the elements of an Excel file, as well as implementations of each interface for both file formats.
When working with the newer `.xlsx` file format, you would use the *XSSFWorkbook*, *XSSFSheet*, *XSSFRow* and *XSSFCell* classes.

In common business scenarios, you need to write a list of business entities in an Excel file.
The first row is reserved for column headers and, the list of business entities, comes in the next rows.

| Id | Name | Quantity | Unit | Price | Currency | Expiration Date |
|----|------|----------|------|-------|----------|-----------------|
| 1  | Beer | 0.33     | L    | 0.50  | €        | 2024-12-31      |
| 2  | Milk | 1.0      | L    | 0.65  | €        | 2022-12-31      |

Another common scenario is the reverse operation. Sometimes, it is necessary to import data into the system from an Excel file.
In this case, if you follow the previous template, you just need to exclude the first row and read the remaining rows.

## Project structure

This project was generated with [Spring Initializr](https://start.spring.io/).
All of the app's code goes in a folder named `src/main`.
The **unit tests** and **integration tests** are in the `src/test` and `src/integrationTest` folders.
Static files are placed in `src/main/resources` folder.

## Available gradle tasks

The tasks in [build.gradle](build.gradle) file were built with simplicity in mind to automate as much repetitive tasks as possible and help developers focus on what really matters.

The next tasks should be executed in a console inside the root directory:

- `./gradlew tasks` - Displays the tasks runnable from root project 'app'.
- `./gradlew bootRun` - Runs this project as a Spring Boot application.
- `./gradlew check` - Runs all checks.
- `./gradlew test` - Runs the unit tests.
- `./gradlew integrationTest` - Run the integration tests.
- `./gradlew clean` - Deletes the build directory.
- `./gradlew build` - Assembles and tests this project.
- `./gradlew help` - Displays a help message.

For more details, read the [Command-Line Interface](https://docs.gradle.org/current/userguide/command_line_interface.html) documentation in the [Gradle User Manual](https://docs.gradle.org/current/userguide/userguide.html).

## Running unit tests

Unit tests are responsible for testing of individual methods or classes by supplying input and making sure the output is as expected.

Use `./gradlew test` to execute the unit tests via [JUnit 5](https://junit.org/junit5/), [Mockito](https://site.mockito.org/) and [AssertJ](https://assertj.github.io/doc/).
Use `./gradlew test -t` to keep executing unit tests in real time while watching for file changes in the background.
You can see the HTML report opening the [index.html](build/reports/tests/test/index.html) file in your web browser.

It's a common requirement to run subsets of a test suite, such as when you're fixing a bug or developing a new test case.
Gradle provides different mechanisms.
For example, the following command lines run either all or exactly one of the tests in the `SomeTestClass` test case:

```bash
./gradlew test --tests SomeTestClass
```

For more details, you can see the [Test filtering](https://docs.gradle.org/current/userguide/java_testing.html#test_filtering) section of the Gradle documentation.

This project uses [JaCoCo](https://www.eclemma.org/jacoco/) which provides code coverage metrics for Java.
The minimum code coverage is set to 80%.
You can see the HTML coverage report opening the [index.html](build/reports/jacoco/test/html/index.html) file in your web browser.

## Running integration tests

Integration tests determine if independently developed units of software work correctly when they are connected to each other.

Use `./gradlew integrationTest` to execute the integration tests via [JUnit 5](https://junit.org/junit5/), [Mockito](https://site.mockito.org/) and [AssertJ](https://assertj.github.io/doc/).
Use `./gradlew integrationTest -t` to keep executing your tests while watching for file changes in the background.
You can see the HTML report opening the [index.html](build/reports/tests/integrationTest/index.html) file in your web browser.

Like unit tests, you can also run subsets of a test suite.
See the [Test filtering](https://docs.gradle.org/current/userguide/java_testing.html#test_filtering) section of the Gradle documentation.

## Debugging

You can debug the source code, add breakpoints, inspect variables and view the application's call stack.
Also, you can use the IDE for debugging the source code, unit and integration tests.
You can customize the [log verbosity](https://docs.gradle.org/current/userguide/logging.html#logging) of gradle tasks using the `-i` or `--info` flag.

This project includes [Swagger](https://swagger.io/). To get a visual representation of the interface and send requests for testing purposes go to <http://localhost:8080/swagger-ui.html>.

## Reference documentation

For further reference, please consider the following articles:

- [Official Gradle documentation](https://docs.gradle.org)
- [Spring Boot Gradle Plugin Reference Guide](https://docs.spring.io/spring-boot/docs/2.5.5/gradle-plugin/reference/html/)
- [Working with Microsoft Excel in Java](https://www.baeldung.com/java-microsoft-excel)
- [Spring Boot File Upload / Download Rest API](https://www.callicoder.com/spring-boot-file-upload-download-rest-api-example/)
- [Testing a Spring Multipart POST Request](https://www.baeldung.com/spring-multipart-post-request-test)
- [Mockito Tutorial](https://www.baeldung.com/mockito-series)
