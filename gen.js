var outlook = Application('Microsoft Outlook')
var app = Application.currentApplication()
app.includeStandardAdditions = true

function titleCase(str) {
    str = str.toLowerCase();
	str = str.split(' ');
	for (var i = 0; i < str.length; i++) {
        str[i] = str[i].charAt(0).toUpperCase() + str[i].slice(1);
	}
	return str.join(' ');
}

var name = app.displayDialog("What's your name?", {
    defaultAnswer: "",
    withIcon: "note",
    buttons: ["Cancel", "Continue"],
    defaultButton: "Continue"
})

var title = app.displayDialog("What's your Job Title? (Remember to be Case Sensitive)", {
    defaultAnswer: "",
    withIcon: "note",
    buttons: ["Cancel", "Continue"],
    defaultButton: "Continue"
})

var phone = app.displayDialog("What's your Phone Number?", {
    defaultAnswer: "01344 623714",
    withIcon: "note",
    buttons: ["Cancel", "Continue"],
    defaultButton: "Continue"
})

template = `
<br />
<table width="100%" style="border-spacing: 0px">
    <tr>
        <td>
            <a href="https://bink.com/">
            <img alt='Bink' title='Bink' width="250" style="margin-bottom:30px" src="https://api.gb.bink.com/content/logos/bink-mono-teal_resize_2023.png">
            </a>
        </td>
    </tr>
    
    <tr>
        <td style="font-family: Arial, Helvetica, sans-serif; color: #30AFB9; font-size: 26px; font-weight: 500; white-space: nowrap;">#name#</td>
    </tr>
    
    <tr>
        <td style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; letter-spacing: 0.8 px; color: #30AFB9; padding-bottom: 10px; white-space: nowrap;">#title#</td>
    </tr>
    
    <tr>
        <td style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #30AFB9; padding-bottom: 5px;">#phone#</td>
    </tr>
    
    <tr>
        <td style="font-size: 14px; font-family: Arial, Helvetica, sans-serif;  padding-bottom: 10px;">
            <a href="https://bink.com/" style="color: #30AFB9; text-decoration: none">bink.com</a>
        </td>
    </tr>
    
    <tr>
        <td width="100%" colspan="2"  style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; padding: 0px; padding-top: 5px;">
            <a href="https://uk.linkedin.com/company/bink" style="text-decoration:none; font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #30AFB9;">
                <img alt='Linkedin' title='Linkedin' width="30" src="https://api.gb.bink.com/content/logos/linkdn_2023.png">
            </a>&nbsp;
            <a href="https://www.facebook.com/BinkLoyalty" style="text-decoration:none; font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #30AFB9;">
                <img alt='Facebook' title='Facebook' width="30" src="https://api.gb.bink.com/content/logos/fb_2023.png">
            </a>&nbsp;
            <a href="https://www.instagram.com/bink_loyalty" style="text-decoration:none; font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #30AFB9;">
                <img alt='Instagram' title='Instagram' width="30" src="https://api.gb.bink.com/content/logos/insta_2023.png">
            </a>
        </td>
    </tr>

</table>
`

var html = template.replace("#name#", titleCase(name.textReturned))
                   .replace("#title#", title.textReturned)
                   .replace("#phone#", phone.textReturned)

sig = outlook.Signature({
    name: "Bink - " + titleCase(name.textReturned),
	content: html
})

sig.make()
