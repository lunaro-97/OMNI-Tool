OMNI is a customer tool which consumes data from standard templates.

In Ericsson, we use 4 different files to match operational changes with some required inputs from BR VIVO customer in Telecommunication.

- base_edb.xlsx: Contains the current project data stored in Ericsson tools and databases
- base_omni.xlsx: Contains the actual version information stored in customer tool
- colunas_dependentes.xlsx: Dependency matrix to build some column filter logic
- de-para_colunas: Fields correlation between Ericsson and customer project data

Using the mentioned code, it compares both databases, clean incorrect/missing values and generate 2 files as output:

- OMNI Upload file: File that contains all newest/updated data resulted from code with proper field formatting
- OMNI Log: List of changes made to original dataframe from customer tool to do some cross-checks and store for future queries
