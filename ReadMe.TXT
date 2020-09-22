Greetings, fellow VB programmer!

This is an add-in for Visual Basic 6.0 that creates code to help you quickly build VB code for ADO or DAO Recordsets and/or SQL statements from your existing Microsoft Access database tables and queries. I created this because I had tried to find a quick, easy and inexpensive tool like this for years, but never had any luck. I previously submitted the original code to my favorite Visual Basic Web site, Planet Source Code (http://www.planetsourcecode.com/) and received a great deal of very positive feeback from other VB programmers, so now I'm also releasing this separately to the VB programming community as a VB 6 add-in, for a very small shareware fee ($20). The shareware product release is called VBRSGen, and is available from my Web site, at http://www.battleszone.com/vbrsgen/ and some other shareware sites.

BUT THIS COPY IS FREE?

Yep, you, however, as a member of the PSC community and fellow contributor of your own source code, get to download this and use it for free! Source code and all! But I'm only submitting it for other programmers' use and feedback on PSC; please do not repost this anywhere else.

I hope you enjoy it and find that it does save you a lot of time typing and retyping tedious code.

I've built dozens of applications that use Access databases as their back-end. I usually keep a copy of the appropriate version of Access running on my workstation along with VB so that I can switch back and forth from VB to the Access interface to cut and paste SQL statements from queries I've designed and saved, or to look at the tables to keep track of which fields and values are where. BBRSGen saves a lot of time, because you can simply point it to your Access MDB file, select the tables or queries you want to work with, and get most of the necessary items (and the correct spellings) ready to paste and use.

VBRSGen creates somewhat tidily formatted code that you can paste into a form or code module in your existing VB project, with a small amount of cleaning up and editing necessary to make the generated code functional for your needs. 

 *=*=*=*=*=*=*=*=*=*
 HOW TO USE VBRSGEN
 *=*=*=*=*=*=*=*=*=*

Getting Started
===============

Open the project file BBRSGEn.VBP in Visual Basic 6.0, and then compile it as a VB Add-In (.DLL), as VBRSGen.DLL, and shut down VB. The next time you open Visual Basic 6.0, you should see VBRSGen right on your list of items under the Add-Ins menu item. The add-in program itself is a file called VBRSGen.DLL, which VB's compile process should have Registered for you (Or you could manually run regsvr32.exe on VBRSGen.DLL to get it into your PC's Registry yourself, but that shouldn't be necessary.)

What Does it Do?
================
VBRSGen will attempt to transform any valid Microsoft Access database table or query into a properly formatted block of code to create a Recordset, including creating and initializing variables for ADO or DAO data access types. Because of the huge variety of possible SQL statements, it's almost impossible for the program to accurately parse every clause exactly as you might expect, so always test the code and "tweak" it by hand, if necessary, to ensure that it provides the results you need.

You can select Microsoft Active Data Objects (ADO) or Data Access Objects (DAO) as your data access model, simply select the proper one from the check box or the Options menu.

VBRSGen will create a connection string to the database you select, but if you need to allow users of your program to find a database in a different path, be sure to change the VB code accordingly.

VBRSGen can create VB code to open an ODBC connection to your Access database using Jet (DAO) or ADO, but to run the code in your project, you MUST set the appropriate reference(s) in your VB IDE. Select Project > References from the VB menu and click on Microsoft Data Access Objects (DAO) or Microsoft Active Data Objects (ADO), as required. If there selections are not available, you may need to download and install the latest release of Microsoft Data Access Components (MDAC). You can download a file called mdac_typ.exe from http://www.microsoft.com/data/download.htm

To use VBRSGen, select VBRSGen from the Visual Basic Add-Ins menu. Then on the main screen, select an Access database file (.mdb). Click on one of the options, Tables or Queries, and then choose the object you want to use from the list. As long as you're selecting a table or a SELECT query, you can even display some of your data in a grid, just by clicking the View in Grid button. Obviously, you can't display an INSERT, UPDATE or DELETE query in a grid, and some complex SQL syntax or queries that require parameters being passed in won't display in the grid either.

Now you have 2 choices: (1) If you select a Query, you can then press the Format SQL button to go straight to the SQL generation screen. (2) If you choose a Table, the next thing to do is select the Fields you want to include (hold down Ctrl while clicking with the mouse to select multiple fields). On the next screen you'll see the fields you selected, and if you want to create query criteria, select a field and then a logical operator (ie, =, >=, <>, etc) followed by a value. You may add further criteria by clicking AND or OR and then selecting another field, logical operator, value, etc. When you're done creating the criteria, click OK and you'll come to the code generation screen. At the code generation screen, you simply pick any options you want to set, and then press Generate Code to have the VB code created in the box below. Choose the output option you prefer and then press OK to copy the code to the Clipboard or into a text file in Notepad, which you can then paste into your project (or save as a text file).

LIMITATIONS AND ISSUES
======================
BE AWARE: This is only version 1.0.0 and there's a lot more I plan to do to make this program even better. For example, sometimes the program gets a bit lost trying to properly parse and format some queries. You should still wind up with enough code to get you well on your way, but never trust the code that's pasted to be error-free, valid SQL syntax. As I mentioned, this is a tool to save you some time, but it's not intended to be a full-blown, click-and-go application generator.

And in case you want to run this as a regular, integral part of your arsenal of VB Add-Ins and you DO want to voluntarily donate to the shareware cause...

You can use any regular credit card by going through PayPal. All you have to do is go to my Web site at http://www.battleszone.com/vbresgen/ and click on the PayPal button. If you're already signed up with PayPal you'll be done in seconds. If not, it'll offer to sign you up--all that means is that you authorize them to charge the credit card of your choice if you ever want to buy things online from anyone else who accepts PayPal payments. I started with using PayPal a couple of years ago for buy stuff on eBay auctions, and they have been great to deal with, no problems for me.

If you can't or don't want to use PayPal, you can mail me a check or money order. It takes longer, but as long as your check is good, that's fine, too. Make out your check to me personally, Brian Battles, and send it to: 

PO Box 310203
Newington, CT 
06131-0203

I'm not charging for sales tax, and I'm not offering disks or CD-ROMs, so there's no shipping charges.

PLEASE PLEASE PLEASE make sure you send along your name, mailing address, telephone number and e-mail address! Otherwise I won't know who's registered and how to reach them!

If you do decide to voluntarily pay $20 I'll add you to the registered customers database and if you register before the end of 2002, I'll place you on my list of permanent customers, meaning that you can have every subsequent release of this program at no further charge. There's a long list of additional features, tweaks and enhancements coming along, including better handling and parsing of SQL statements, possible support for other database products (ie, ODBC data sources, SQL Server, delimited text files, Excel sheets, Oracle, anything with a valid DSN, etc). Whenever I update this, the latest and greatest version will be available at www.battleszone.com, and registered users will always get new versions free.

Looking to the future, if this ever grows into anything far more complex and advanced, it will probably get a whole new name and new look, and then it would be released as a whole different product, but I'll still provide great discounts for you who paid for copies of this product.

I'm also known to be kind of a softie when it comes to giving assistance to children, fellow Masons and ham radio operators...speak up! 


SUPPORT
=======
As far as support is concerned, I'm not running a big software company with a brigade of customer service techs answering telephones 24 hours a day, so go easy on me! I'm a fairly nice guy, and I'm going to try to do whatever it takes to make you happy. If something acts really strange or won't work on your machine, shoot me an e-mail or call me on the phone and I'll see what I can do. But let's face it, I can't fix every imaginable problem, and it isn't worth spending dozens of hours struggling with a simple utility like this if there are serious problems on your end. So don't expect miracles, but enjoy having VBRSGen help cut down some of the time and tedium of routine database coding!

CONTACT
=======
I would love your feedback, comments, suggestions, and even a few complaints if you aren't too crabby. Drop me a note to brianb@cmtelephone.com or call me at 860-808-1285. (I'm in Connecticut, so you may not catch me in my office before 9 AM or after 6 PM Eastern US Time.) Visit my Web site at http://www.battleszone.com/ from time to time and see if I've posted a newer version or any other nifty programs, news or other useful stuff.

Thanks for downloading VBRSGen and bothering to actually read this file! Please register and enjoy using this VB tool.

Have an exceptional day!

  --Brian