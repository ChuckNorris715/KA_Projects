###################################################################################################
##	This AppleScript is for generating Mail Signature Templates in Microsoft Outlook for Mac	 ##
##	Prepared by Stuart Lamont in November 2015 to replace the Centenary Signatures				 ##
##  																							 ##
##	Modified to use Enterprise Connect to preform numerous Active Directory Lookups 			 ##
##  and will produce inconsistent results if the Active Directory binding is in 				 ##
##  any way compromised. If Surname, Title and Phone Number aren't on the Generated 			 ##
##  Template, the computer with Active Directory and the re-run the script.						 ##
##																								 ##
##	If the Script is run more than once, multiple Templates will be generated, so please		 ##
##	bear this in mind when selecting the default templates for the user.						 ##
###################################################################################################

#set MyName to name of me as string
#display dialog MyName
#instantiate global variables
#global variables are used here for the subroutine to access
global longName
global userName
global rawsurname
global firstname
global surname
global credentials
global email
global jobTitle
global phoneNo
global directPhone
global address1
global descript1
global descript2
global fontColour1
global fontColour2
global location1Name
global location2Name
global descriptMain

#Variables for Graphics Assets
global logoLink
global webURL
global webURLText
global twitterLink
global twitterLogoLink
global facebookLink
global facebookLogoLink
global linkedInLink
global linkedInLogoLink
global instaLink
global instaLogoLink
global bottomBorderImage

#variable for HTML Block
global HTMLContent


#Collect User Data and place in Variable containers
tell (get system info)
	set longName to long user name
	set userName to short user name
end tell

#pull Surname from System LongName
#not used
set rawsurname to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a sAMAccountName | awk 'BEGIN {FS=\": \"} {print $2}'")
#pull first name from System LongName
set firstname to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a givenName | awk 'BEGIN {FS=\": \"} {print $2}'")
#Pull Surname from AD Attribute surname 
set surname to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a sn | awk 'BEGIN {FS=\": \"} {print $2}'")
#pull pull "credentials" (LEED/RA/AP) from AD Attribute description 
set credentials to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a description | awk 'BEGIN {FS=\": \"} {print $2}'")
#pull email address from AD Attribute mail
set email to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a mail | awk 'BEGIN {FS=\": \"} {print $2}'")
#pull job title from AD Attribute title
set jobTitle to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a title | awk 'BEGIN {FS=\": \"} {print $2}'")
#pull telephone Extension number from AD Attribute telephoneNumber (i.e. extention)
set phoneNo to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a telephoneNumber | awk 'BEGIN {FS=\": \"} {print $2}'")
#pull direct telephone from AD Attribute ipPhone (i.e. direct number)
set directPhone to do shell script ("\"/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl\" -a ipPhone | awk 'BEGIN {FS=\": \"} {print $2}'")

#####################################################
#Setup Addresses
set address1 to "756 Haddon Ave. Collingswood, NJ 08108"


#####################################################
#Setup to add a credentials line
set descript1 to "<tr>
                 <td valign=\"top\" align=\"left\" class=\"qe_defaultlink\" style=\"font-family: 'Montserrat', Arial, sans-serif;font-size:11px;line-height:15px;color:#231f20; font-weight:600;padding-top:6px;\"> " & credentials & "</td>
               </tr>"
set descript2 to ""

#####################################################
#Setup Location Names
#set location1Name to "Ivanhoe"
#set location2Name to "Plenty"

#####################################################
#setup graphical Assets
set logoLink to "http://zerozone.com/qeinbox/signatures/logo_updated.png"
set webURL to "https://www.thriven.design/"
set webURLText to "thriven.design"
set facebookLink to "https://www.facebook.com/thriven.design/"
set facebookLogoLink to "http://zerozone.com/qeinbox/signatures/facebook.png"
set linkedInLink to "https://www.linkedin.com/company/thrivendesign"
set linkedInLogoLink to "http://zerozone.com/qeinbox/signatures/linkedin.png"
set instaLink to "https://www.instagram.com/thriven.design/"
set instaLogoLink to "http://zerozone.com/qeinbox/signatures/insta.png"
#set bottomBorderImage to "http://media.igs.vic.edu.au/general/signatures/bottomborder.jpg"

#Error Checking
#check for field data complete - If surname is Blank, quit, and prompt user to come to IT Services
#if surname is "" then
#	display dialog "This Action cannot be completed as your computer's Active Directory Binding is broken. Please bring your computer to IT Services to correct this issue." with icon stop buttons "Exit"
#	return

#end if



###############################################################################################################################################################

#Prompt user to select which Campus they are based at. Will determine which template is generated
#set question to display dialog "Which Campus are you based at?" buttons {location1Name, location2Name} default button location1Name
#set campus to button returned of question
#Ridgeway Template
if credentials is equal to "" then
	
	set descriptMain to descript2
	setupSignature()
	
else
	
	set descriptMain to descript1
	setupSignature()
	
	
end if

on setupSignature()
	tell application id "com.microsoft.Outlook"
		make new signature with properties {name:"Thriven_Signature", content:"<html>
<body class=\"qe_body\" style=\"padding:0; margin:0 auto !important; display:block !important; min-width:100% !important; width:100% !important; background:#ffffff; -webkit-text-size-adjust:none\">
<table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" bgcolor=\"#ffffff\"  class=\"full-wrap\">
  <tr>
    <td align=\"center\" valign=\"top\"><table align=\"left\" style=\"width:325px; max-width:325px; table-layout:fixed;\" class=\"qe_wrapper\"  width=\"325\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">
        <tr>
          <td valign=\"top\" align=\"center\" style=\"padding:20px 6px;\"><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" align=\"center\">
              <tr>
                <td valign=\"middle\" align=\"left\" width=\"104\" style=\"width:104px;padding-top:4px;\"><a href=\"" & webURL & "\" target=\"_blank\" style=\"text-decoration:none;\"><img src=\""& logoLink &"\" width=\"104\" alt=\"thriven design\" border=\"0\" style=\"font-family:Arial, sans-serif; font-size:14px; line-height:17px;color:#000000;display:block;max-width:104px;\"/></a></td>
                <td valign=\"middle\" align=\"center\" style=\"padding-left:15px;\"><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">
                    <tr>
                      <td valign=\"top\" align=\"left\" class=\"qe_defaultlink\" style=\"font-family: 'Montserrat', Arial, sans-serif;font-size:16px;line-height:20px;color:#231f20; font-weight:bold;\">" & firstname & "&nbsp;" & surname & "</td>
                    </tr>
				" & descriptMain & "
                    <tr>
                      <td valign=\"top\" align=\"left\" class=\"qe_defaultlink\" style=\"font-family: 'Montserrat', Arial, sans-serif;font-size:10px;line-height:13px;color:#000000; padding-top:5px;\">" & jobTitle & "</td>
                    </tr>
                    <tr>
                      <td valign=\"top\" align=\"center\" style=\"padding:5px 0px;\"><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" align=\"left\" >
                          <tr>
                            <td height=\"1\" style=\"height:1px;font-size:0px;line-height:0px;\" bgcolor=\"#000000\"></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr>
                      <td valign=\"top\" align=\"left\" class=\"qe_defaultlink\" style=\"font-family: 'Montserrat', Arial, sans-serif;font-size:9px;line-height:13px;color:#000000;\"><a href=\"mailto:" & rawsurname & "@thriven.design\" style=\"text-decoration:none;color:#000000;\">" & rawsurname & "@thriven.design</a></td>
                    </tr>
                    <tr>
                      <td valign=\"top\" align=\"left\" class=\"qe_defaultlink\" style=\"font-family: 'Montserrat', Arial, sans-serif;font-size:9px;line-height:13px;color:#000000;padding-top:5px; \"><strong>t: </strong><a href=\"tel:" & phoneNo & "\" style=\"text-decoration:none;color:#000000;\">" & phoneNo & "</a>&nbsp;|&nbsp;<strong>d: </strong><a href=\"tel:" & directPhone & "\" style=\"text-decoration:none;color:#000000;\">" & directPhone & "</a></td>
                    </tr>
                    <tr>
                      <td valign=\"top\" align=\"left\" class=\"qe_defaultlink\" style=\"font-family: 'Montserrat', Arial, sans-serif;font-size:9px;line-height:13px;color:#000000; padding-top:5px;\">" & address1 & "</td>
                    </tr>
                    <tr>
                      <td valign=\"top\" align=\"center\" style=\"padding-top:6px;\"><table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" align=\"center\">
                          <tr>
                            <td valign=\"top\" align=\"left\" width=\"26\" style=\"width:26px; line-height:0px; font-size:0px;\"><a href=\"" & linkedInLink & "\" target=\"_blank\" style=\"text-decoration:none;\"><img src=\"" & linkedInLogoLink & "\" width=\"22\"  border=\"0\" style=\"font-family:Arial, sans-serif; font-size:14px; line-height:17px;color:#000000;display:block;max-width:22px;\"/></a></td>
                            <td width=\"5\" style=\"width:5px;line-height:0px;font-size:0px;\"></td>
                            <td valign=\"top\" align=\"left\" width=\"27\" style=\"width:27px; line-height:0px; font-size:0px;\"><a href=\"" & instaLink & "\" target=\"_blank\" style=\"text-decoration:none;\"><img src=\"" & instaLogoLink & "\" width=\"22\"  border=\"0\" style=\"font-family:Arial, sans-serif; font-size:14px; line-height:17px;color:#000000;display:block;max-width:22px;\"/></a></td>
                            <td width=\"5\" style=\"width:5px;line-height:0px;font-size:0px;\"></td>
                            <td valign=\"top\" align=\"left\" width=\"27\" style=\"width:27px; line-height:0px; font-size:0px;\"><a href=\"" & facebookLink & "\" target=\"_blank\" style=\"text-decoration:none;\"><img src=\"" & facebookLogoLink & "\" width=\"22\"  border=\"0\" style=\"font-family:Arial, sans-serif; font-size:14px; line-height:17px;color:#000000;display:block;max-width:22px;\"/></a></td>
                            <td width=\"10\" style=\"width:10px;\"></td>
                            <td valign=\"middle\" align=\"left\" class=\"qe_defaultlink\" style=\"font-family: 'Montserrat', Arial, sans-serif;font-size:10px;line-height:13px;color:#000000;font-weight:600; \"><a href=\"" & webURL & "\" target=\"_blank\" style=\"text-decoration:none;color:#000000;\">" & webURLText & "</a></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>"}
	end tell
	
end setupSignature


##########################################################
#This subRoutine currrently will not function without "Accessibility Access" enabled for the app.
on updateDefaultSig to mySignature for accountName
	tell application "Microsoft Outlook"
		activate
	end tell
	
	tell application "System Events"
		tell process "Microsoft Outlook"
			tell menu bar 1
				tell menu bar item "Outlook"
					tell menu "Outlook"
						click menu item "Preferences..."
					end tell
				end tell
			end tell
		end tell
	end tell
	
	tell application "System Events"
		tell process "Microsoft Outlook"
			click button 7 of window "Outlook Preferences"
		end tell
	end tell
	
	tell application "System Events"
		tell process "Microsoft Outlook"
			tell window "Signatures"
				tell group 2
					---click pop up button 2
					set Preset to get value of pop up button 2
					if Preset is equal to "Thriven_Signature" then
						
					else
						if Preset is equal to "None" then
							click pop up button 2
							delay 0.5
							keystroke (ASCII character 31) -- down arrow key 
							keystroke (ASCII character 31) -- down arrow key 
							delay 0.5
							keystroke (ASCII character 3) -- enter key
							delay 0.5
							
						else
							click pop up button 2
							delay 0.5
							keystroke (ASCII character 31) -- down arrow key
							delay 0.5
							keystroke (ASCII character 3) -- enter key
							delay 0.5
						end if
					end if
					set Preset to get value of pop up button 1
					if Preset is equal to "Thriven_Signature" then
						
					else
						if Preset is equal to "None" then
							click pop up button 1
							delay 0.5
							keystroke (ASCII character 31) -- down arrow key 
							keystroke (ASCII character 31) -- down arrow key 
							delay 0.5
							keystroke (ASCII character 3) -- enter key
							delay 0.5
							
						else
							click pop up button 1
							delay 0.5
							keystroke (ASCII character 31) -- up arrow key
							delay 0.5
							keystroke (ASCII character 3) -- enter key
							delay 0.5
						end if
					end if
				end tell
			end tell
		end tell
	end tell
end updateDefaultSig