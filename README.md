# district

This script matches an address with a voting district.
Input: exel file with two sheets. Sheet 1 contains addresses for matching (column 1) as strings and their districts to be inserted (column 2). Sheet 2 contains distribution of addresses between districts.
The idea of the solution is to parse an address to get a street name and a building number. Sheet 2 forms an array of objects with which streets and buildings are compared. If number of building is not specified and the street belongs to several districts, then the district is undefined. The same happens if only city is specified.
