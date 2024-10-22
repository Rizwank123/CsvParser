### Theory Questions ###
How will you tackle the challenge above?

I would first parse the CSV input and load it into a data structure like a 2D array. For each cell, if it contains a formula, I will evaluate it by resolving the referenced cells and performing the required operation. Finally, I would generate a new CSV file with the computed values.

What type of errors would you check for?

I would check for:

    Invalid or unrecognized formulas.
    Circular references between cells.
    Division by zero errors.
    Missing or out-of-bounds cell references.
    Non-numeric data where numeric values are expected.

How might a user break your code?

A user could break the code by:

    Providing malformed formulas.
    Introducing circular dependencies between formulas.
    Referencing cells that don't exist.
    Using operations that result in errors (e.g., division by zero).
    Entering unexpected data types (e.g., text in numeric fields).
