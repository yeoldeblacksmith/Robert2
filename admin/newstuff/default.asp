<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>What's New</title>
        
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "What's New"

    navmenu.WriteNavigationSection ""
%>
        <table class="adminTable" style="margin: 0 auto; width: 1050px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <table class="tableDefault marginCenter" style="border-spacing: 5px">
                        <tr>
                        <td>
                          
 <a href="https://www.vantora.com/videos.asp#PayPalSetup" target="_blank"><img src="http://www.vantora.com/images/videohelp.gif" border="0" align="right" alt="Setting up Paypal" title="Setting up Paypal Video"></a>   
 <h2>Jan 7 / 14 - Setting up Paypal</h2>
 
You can now accept deposits, payments for add ons, etc by easily setting up a paypal account and making a few quick modifications to your admin pages.  You can set over all deposits, or monitize custom fields where you can charge for different options. The video will walk you through getting going!                     
                        
                        
   <h2>Dec 25 - Custom Fields on Waivers</h2>                      
                        
 <a href="https://www.vantora.com/videos.asp#CustomFieldsInWaivers" target="_blank"><img src="http://www.vantora.com/images/videohelp.gif" border="0" align="right" alt="Custom Fields on Waivers Video" title="Custom Fields on Waivers Video"></a>   

 
Had a request to put custom fields on the waivers, so we did.  It's just like building the custom fields on registration, except you can choose if they go to a Generic area, the Adult area, the Minor area, or Minor and Adult.  For example, maybe on each waiver you need some additional info for each guest, like would they want to rent, do they want a pizza slice... anything you want to add you can.  


 <h2>Nov 1 - Total Max Players at your Facility</h2> <a href="https://www.vantora.com/videos.asp#maxplayers" target="_blank"><img src="http://www.vantora.com/images/videohelp.gif" border="0" align="right" alt="Custom Fields on Waivers Video" title="Setting Max Players Help Video"></a>   
We added a total Max Players field in your Admin -> Settings -> Registration page.  We defaulted it to 1000.  This number is the max that can be booked in at any given time.  So, for example, if your time is set for 2 1/2 hours, and you have your max set at 100, and 100 people are in at noon, it won't let anyone book until 2:30 in the afternoon.  It also won't let people book in after 9:30 that day, knowing there are 100 at noon.  This is useful if you have a situation, where perhaps you only have 100 sets of gear.  It would not be good then to have 140 people wanting to play.
<br><br>
To make this work, we have also changed the look of the registration form slightly.  Now it asks for the customers name, email, and number of people in their party.  Then after they enter the number of the party, it then will calculate what available times exist that will allow a party of that size to book in.

                        
<h2>Oct 15 - Big Changes to Registration Custom Field blurbs</h2>   
We have made some big changes to the custom field blurbs.  Before, anything you put in the blurb area for a field ended up showing on the Registration form all the time, which made the form very confusing and cluttered looking.  Now, instead we have the blurb showing up as a pop up "tool tip" which is content sensitive.  That is to say, when a customer puts their mouse over the times, the tool tip will pop up to explain the time.  Or if you have a custom field, with a blurb, it's once the customer mouses over that field that the tool tip will show up.  Then it will go away when moving off that field.
<br><br>
Now it turns out people wanted even more infomation on the registration, so we have also added an additional area for "Additional Information".  This allows you to put a lot of info that will pop up in a window only when someone clicks on the "info" button that will automatically appear to the right of the field if you  put additional information in there.  If not, the button will not show up.  You can use full html, including adding photos, etc to the additional information area.
<br><br>
To see an example of the changes, feel free to go to my site, <a href="http://www.vantora.com/paintball/gatsplat/registration/" target="_blank">GatSplat Reservations at Vantora</a> then pick any date and check out the 50 caliber option we have near the bottom of the registration form. 
<br><br>
If you need any help utilizing this new feature, just like any other feature, don't hesitate to contact me and I can walk you through it, or help you get it going.

                    
                        
<h2>Sep 25 - Removed Event Type</h2>
So, what we did was remove the "Event Type" in your registration screen.  But we didn't lose your information there.  We copied everyone Event Type info into a "Custom Field".  The reason we did this, is we are getting ready to add additional functionality to the Custom Field area.  For example, some asked for the selection options have the ability of being moved up or down.  That will be available soon.  The other thing, is we are tying a pay pal and deposit ability to custom fields in the next few weeks.
<br>
<br>
For example, I was just using "Event Type" for things like, "Birthday Party", "Corporate Event", "School/Youth Function", etc, but I saw others using it for things like, "Bronze Package", "Silver Package", "Gold Package", etc.  And we know that some people want the ability to take deposits for their reservations.  So in the near future, you will be able to go in, and as an example, choose to put a $100 deposit on a "Bronze Package", or a $150 deposit on a "Silver Package", etc.  Or you will be able to choose a $10 per person deposit on Bronze Package, so a group of 6 choosing bronze package would have a $60 deposit, etc.
<br><br>
Since the "Event Type" was really just a watered down version of a custom field, we decided just removing it and moving your data to a custom field made more sense.  And now you can change the title of that custom field if you wish.  So instead of "Event Type" you could go name it "Package Options" etc.



<h2>Sep 19 - Email Previews</h2>
As you know, we default the main info in the outgoing emails for you.  But you can add whatever you want to them, using the editor.  Add a note, a special, etc.  Now, after you are done, you can hit the "Preview" button found below the editor, next to the save button, and it will pop up a new window and show you what the outgoing email will look like.  Then you can make changes, save and preview again.  Since it is a pop up window, you will want to remember to close it after viewing.  If you click preview a second time and it seems like nothing happened, the preview window is probably already open, so just look for it on your task bar, or hit Ctrl Tab.

<h2>Aug 21 - Bug Fix Waiver CDate error</h2>
<p>Occasionally, on either the normal, or the kiosk version of the waiver, it would crash giving a page with a CDate error.  But it was very infrequent, and we did not know how to duplicate it.  Finally it came to me, it was an error that occurred when people would make up illegal dates on the form.  So if a customer just picks numbers, it was possible in the old format for them to choose the month, say February, then pick a day, and 1-31 was available in the drop down.  If they would just go to the bottom and pick the 31st of February in some year, obviously this is not an actual date.  This would cause the software to show the error when trying to translate that into a real date.</p>
<p>This has been corrected by making the customer pick the year first, then the month, then the day drop down will only allow legal dates for that particular month (and year in the case of leap years).</p>

<h2>July 25 - Waiver Downloads</h2>
<p>For those who want to download your waiver info, we have created a module that lets you do that.  You will see it if you mouse over Waivers in your admin menu, then choose export. You can choose to download all of them, the month, the year, or between custom dates.  When you download waivers, they are marked as downloaded, so next time you do the download, you don't have to bring down the ones you already have, creating duplicates.  But if you do want them all again, there is a box to check for download all, including previously downloaded waivers.</p>
<p>The format is a delimited CSV file which will be named with your field name, then the date options you choose.  This file, if imported into Excel, or any other spread sheet program, will automatically set up the columns for you and import easily.</p>


 <h2>July 3 - Custom Fields for Registrations</h2>
<p> So, we have added the ability for you to make your own custom fields on the registration form.  In the admin section, it is under  Settings -> Registration -> Custom Fields.  They will show up near the bottom, after the other fields, but before the submit.  You have many options when creating these.  They can be Yes / No, or you can choose short text, or a larger memo box type text, or you can pick option which will turn it into a drop down box type field.  There is also the option for making a field required so they will not be able to leave the form until they pick an answer. 
 <br><br>
 You will also see a notes section.  This will show up just below the field so you can add extra notes.  For example, perhaps your question is:  Upgraded Guns?: Yes No   then the notes could be: The upgraded guns have a longer barrel and shoot straighter.  If all members of your group use them, they will only be $3.00 extra for each player.
 <br><br>
 And just like all the blurbs, you can pick the size, color, etc of this note, add images, etc.  I would, however, suggest you keep notes very short, and if you want to explain the difference between packages, or guns, do this on your pages, and maybe just run a link with a note pointing them to additional info.  If setting links using the cute writer, make sure you choose the <b>Target - New Window</b> option so it does not remove them from the page.  Keep these types of things to a minimum as you want to avoid making the registration process too clumsy or cluttered.  

</p><hr>
 <h2>June 26 - Improved Password Retrieval</h2>
 Not too important unless you have a very bad memory, but we've improved the way you reset your password.  If you do it wrong a few times, you can click on the "forgot password?" link, and it will prompt you to put in your email.  Then it will send an email to you with your username, and a link to click.  If you click that link, it will allow you to reset your password.  If you don't click it.. your password will still remain the same, so if you suddenly remember it, and don't want to change it, that's cool.  Not a big thing, but we figured this was one of those necessary features to make it easier for forgetful paintball field owners!
 
                             <p> <hr>                                         
<h2>June 25 - Removed Opt out of Email in waivers</h2>
After watching people on the kiosks, we determined that about 30% were choosing to opt out of the email list by un clicking the box.  We then also decided that if the email data is mainly used to send a free pass on a birthday, no one would probably complain, and we will make sure we have an easy unsubscribe, opt out link in any outgoing emails we generate using this list. 
 <p><hr>
 
 <h2>June 1 - Expanded Event Type Size</h2>                                           
While I was just using event type for things like "Birthday Party", or "Corporate Event", others expressed a desire to put a bit more in there, like perhaps, "Private Party (Call for Reservations)".  The new maximum field length is 50 characters, instead of 20.                                            
                                            
                                            
                             </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
</html>
