# Document Generation Program

## Table of Contents
- [Introduction](#introduction)
- [Requirements](#requirements)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Documentation](#documentation)
- [Contributing](#contributing)
- [License](#license)

## Introduction

The **Ghost App** is a tool designed to assist in the creation of batch record documents specific to my current manufacturing needs. It automates the process of filling in placeholders and calculations for these documents.

## Requirements

To run this program, you need the following software and tools:

- [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)
- [Microsoft Word](https://www.microsoft.com/en-us/microsoft-365/word)
- [.NET Core](https://dotnet.microsoft.com/download)

## Getting Started

1. Clone or download this repository to your local machine.

2. Open the terminal or command prompt and navigate to the project directory.

3. Ensure you have the required dependencies installed (Excel, Word, .NET Core).

4. Build the project:

    ```shell
    dotnet build
    ```

5. Run the program:

    ```shell
    dotnet run
    ```

## Usage

The program performs the following tasks:

1. **User Input**: The program prompts the user to input the bench they are working at.

2. **Document Creation**: It generates Request Forms and Ligation Batch Records, filling in specific information related to ligation experiments.

3. **Reagent Information**: The program retrieves reagent information from Excel worksheets.

4. **Data Replacement**: It replaces placeholders in Word documents with actual data and calculations.

5. **Document Saving**: The modified documents are saved in the user's temporary directory.

## Documentation

The program is written in F# and relies on various functions for document processing and data retrieval. These functions are well-documented in the source code.

## Contributing

If you'd like to contribute to this project, please follow these steps:

1. Fork the repository.

2. Create a new branch for your feature or fix.

3. Make your changes and commit them.

4. Push your changes to your fork.

5. Create a pull request, explaining your changes and their purpose.

## License

This program is open-source and available under the [MIT License]([LICENSE](https://github.com/IvanPedroza/Ghost-App/blob/master/LICENSE.md)https://github.com/IvanPedroza/Ghost-App/blob/master/LICENSE.md).
