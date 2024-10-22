# CSV Formula Evaluator

## Overview

This program processes a CSV file that contains both numeric values and formulas. It evaluates the formulas by referencing other cells and outputs the results to a new CSV file.

## Features

- Reads a CSV input file.
- Evaluates any formulas within the CSV (e.g., `=A1+5`, `=B2*A3`).
- Outputs the results to a new CSV file with all formulas resolved to their final values.

## Example

### Input CSV

```csv
A,B,C
5,3,=5+A1
7,8,=A2+B2
9,=4+5,=C2+B3

Steps

    Clone the project:

    bash

git clone https://github.com/your-repo/csv-formula-evaluator.git

Navigate to the project directory:



cd CSVPARSER

Build the project:
```
mvn clean package
```
Place your input.csv file in the root directory.



