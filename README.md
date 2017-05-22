# Partial-Match
# By Jason McCoy

Removing Duplicates Based on a Partial Match

Excel Macro for Partial Match
The macro shown here simply checks the first X characters of a "key" value against a range and returns the address of the first matching cell.

For example, let's assume that your addresses are in the range A2:A100. In column B you can use this NearMatch function to return addresses of possible duplicates. In cell B2 enter the following formula:

=NearMatch(A2,A3:A$100,12)

The first parameter for the function (A2) is the cell you want to use as your "key." The first 12 characters of this cell are compared against the first 12 characters of each cell in the range A3:A$100. If a cell is found in that range in which the first 12 characters match, then the address of that cell is returned by the function. If no match is located, then the #N/A error is returned. If you copy the formula in B2 down, to cells B3:B100, each corresponding address in column A is compared to all the addresses below it. You end up with a list of possible duplicates in the original list.
