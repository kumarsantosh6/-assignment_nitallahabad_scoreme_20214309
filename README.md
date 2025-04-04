# -assignment_nitallahabad_scoreme_20214309

## Sample Output

For `test6.pdf`, the output Excel will contain:
- Multiple sheets for each page with structured tables
- Proper column headers from the banking statement
- Clean transaction data

For `test3.pdf`, the output will:
- Parse transaction lines into Date, Description, Amount columns
- Handle debit/credit indicators
- Maintain running balance information

## Future Improvements

1. Add command-line interface
2. Support batch processing of multiple PDFs
3. Add visualization of detected tables
4. Improve accuracy for complex layouts
5. Add unit tests

This solution provides a robust way to extract tables from PDFs while meeting all the assignment requirements. The approach combines structured table extraction with pattern matching for semi-structured data, ensuring maximum coverage of different PDF formats.
