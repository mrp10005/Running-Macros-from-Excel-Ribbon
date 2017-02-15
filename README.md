# Running-Macros-from-Excel-Ribbon
No code here but I run almost all my macros from the ribbon and I think it's very useful.
My friend was telling me he runs his by opening up different excel documents and running the macros from going into VBA from there.
I almost fell out of my chair when I heard that and showed him this.

Steps:
1. Click Developer --> Record macro --> Name the macro --> Store the macro in: Personal Macro Workbook
2. Stop Recording
3. View --> Unhide --> PERSONAL
4. Alt F11
5. Select the drop down for VBAProject(Personal.XLSB)
6. Select the drop down for Modules
7. Select Module1 (the macro you recorded should be here, definitely if this is your first time doing this)
8. Now you can write up your macro here
  8a. I prefer having multiple subroutines for my macros.  If you have one subroutine you don't need to do steps 9 and 10.
9. Insert another module in VBAProject(Personal.XlSB)
10. Name this whatever you'd like and have it do:
  Call name_of_subroutine_of_the_first_sub_you_just_made_in_8
  (Note: if you're like me and have multiple sub routines that call others in sequential order this is best)
11. Click Save in the top left
12. Close out of the VBA window
13. In the Personal Workbook: View --> Hide --> This hides the Personal Workbook
14. Close Excel
  14a. A windown will ask if you want to save the Personal Workbook, select yes
15. Open Excel
16. File --> Options --> Customize Ribbon
  16a. Create a new tab --> Create a new group
17. "Choose Commands From" --> Macros --> Select the macro that will call the macro you want to actually run (see step 10)
18. Click Ok and close
19. Now you will have a new Tab in your Ribbon and you can run you macro by selecting the button in the new tab 
  
