The aim of the program is to read and process data from an excel file and use them to change descriptions and reason codes of the items in SAP FBL5N transaction.
How the program reads data from excel
At the beginning it asks about the saving date, it proposes the current one:
<img width="1048" height="309" alt="image" src="https://github.com/user-attachments/assets/0b402d83-8037-4bac-8b39-01316567fdc8" />
It uses the columns highlighted below:
<img width="1766" height="305" alt="image" src="https://github.com/user-attachments/assets/a536256a-a156-4462-8c09-eb5d81fefe04" />

The program skips the lines with ci-joint and avance, speaking precisely all the lines where the detail is among (lower or upper case):
<img width="740" height="39" alt="image" src="https://github.com/user-attachments/assets/0aee3d8c-58c3-4c99-a20a-0a878b8b8ca2" />
<img width="1004" height="141" alt="image" src="https://github.com/user-attachments/assets/23494434-fd3c-4a35-bfdc-d0b3a48d34ba" />
<img width="1090" height="371" alt="image" src="https://github.com/user-attachments/assets/f2ba8fc6-e3f3-421b-a41f-696e61c22442" />
If a detail is among:
 <img width="638" height="34" alt="image" src="https://github.com/user-attachments/assets/4454a49a-d1ea-4edc-9e70-5312b33b9745" />

the program will code the detail as 0 and later try to describe all the items on the account:

If there are inconsistencies in the numbers of customers, it will warn you about this at the beginning:
<img width="671" height="194" alt="image" src="https://github.com/user-attachments/assets/87b20b93-36c2-455d-950c-6e64a415dddd" />
<img width="1090" height="267" alt="image" src="https://github.com/user-attachments/assets/299bf297-557c-41a3-81ea-2ddffa9a38ed" />
The program treats everything that isn’t a number in detail column just as a separation of items’ numbers:
<img width="501" height="196" alt="image" src="https://github.com/user-attachments/assets/33ae3ec5-1130-4649-a0c1-b03e27711d3a" />

For BOEs with the same detail, bank abbreviation and fairly consecutive numbers, it concatenates the references like so:
<img width="1149" height="119" alt="image" src="https://github.com/user-attachments/assets/a02d8702-587d-45d8-a907-3c01c5d514b1" />
<img width="940" height="185" alt="image" src="https://github.com/user-attachments/assets/8a1e6eff-5dc5-48d6-8e69-be887eaf4ed8" />

