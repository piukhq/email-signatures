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
    <td><a href="https://bink.com/"><img alt='Bink' title='Bink' width="150" height="114" src="https://api.gb.bink.com/content/media/logos/bink_with_text_email_signature.png"></a></td>
  </tr>
  <tr>
    <td style="font-family: Arial, Helvetica, sans-serif; color: #8EB1B7; padding: 0px; font-size: 20px; font-weight: bold; white-space: nowrap;">#name#</td>
  </tr>
  <tr>
    <td style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #8EB1B7; padding: 0px; padding-bottom: 5px; white-space: nowrap;">#title#</td>
  </tr>
  <tr>
    <td style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #8EB1B7; padding: 0px;">#phone#</td>
  </tr>
  <tr>
    <td style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; padding: 0px; padding-top: 5px;"><a href="https://bink.com/" style="color: #CCDBDF; text-decoration: none;">bink.com</a></td>
  </tr>
  <tr>
    <td width="100%" colspan="2"  style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; padding: 0px; padding-top: 5px;">
      <p style="color: #8EB1B7;">
        <a href="https://www.linkedin.com/company/bink" style="text-decoration:none; font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #CCDBDF;">LinkedIn</a>&nbsp;&nbsp;|&nbsp;
        <a href="https://www.facebook.com/BinkLoyalty" style="text-decoration:none; font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #CCDBDF;">Facebook</a>&nbsp;&nbsp;|&nbsp;
        <a href="https://www.instagram.com/bink_loyalty/" style="text-decoration:none; font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #CCDBDF;">Instagram</a>
      </p>
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
