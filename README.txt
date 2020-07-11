Validate Vouchers
Fernanda Cetnarovski, July 2020
https://github.com/fergabi17/
*********************************************************

This script validates codes print on vouchers for Smartbox Group.

- Create a folder on your desktop called "VOUCHERS", containing all vouchers to be checked.
- All vouchers must be in pdf format and named as the following example:

	D1071236_UKU20112A2012P

- The script recognizes the country of the voucher by checking the two first letters after "_". They can be PT or UK. If they are anything other than that, the voucher will be considered as a country from PIM.
- Double click on the script and wait until it finishes.
- A csv file will be generated in the same "VOUCHERS" folder, containing the codes found and a validation check for PIM countries.