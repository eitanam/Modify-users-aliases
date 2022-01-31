# Modify-users-aliases

If you are using an hybrid enviroment and you don't have an Exchange on-prem.
This tool will help you modify the aliases and hide /unhide users

*********************************************************************************
This script is provided AS-IS without any warranty to any damage that may occured.
If you are using it it's AT YOUR OWN RISK!
*********************************************************************************

Version 1.0
Inital release

Version 1.1
Check if the AD module is installed

Version 1.2
Check if the Proxyadresses attribute is empty. In case it is, it will add SMTP address based on the UPN
Add an option to modifr the 365 UPN to be identical to the new primary email address

Version 1.3
Modified the user exsitence procedure

Prerequisite

Install ActiveDirectory Powershell Module


How to use it

run the Modify users aliases.ps1

![image](https://user-images.githubusercontent.com/71331120/151763093-ce49c60a-deeb-42a4-bd0a-df7b6e3ec202.png)

Type the user name you want to modify (user name and not display name) and click on Search

You can toggle between Hide / unhide from the address book
Primary email address - change the primary email address of the user
Cahnge the UPN also on the 365 - will change the user login name also in the 365
Add an alias - will add an alias to the user smtp aliases
Delete an alias - will delete an alias from the user smtp aliases

On the result window you will see the user status (aliases, hidden from the GAL, Enabled in the AD and the user UPN)

![image](https://user-images.githubusercontent.com/71331120/151763796-5b6b37b2-9058-4908-9d12-936a93abb05e.png)
