<form class="baseform validate_form" id="enquiry_form" method="post" action="http://www.webfirm.com.au/scripts/mail.asp">
  <div class="hidden-form">
    <input type="hidden" name="Email_To" class="safemail" value="steven dot taddei at webfirm dot com" />
    <input type="hidden" name="Subject" value="Contact Form Subject" />
    <input type="hidden" name="Thanks" value="contact_thankyou.shtml" />
    <input type="hidden" name="URL" value="http://demo.webfirm.com.au" />
  </div>
  <ul>
    <li>
    <label for="Name">Name <em>*</em></label><input id="Name" name="_1_Name" type="text" class="required replace_text" autocomplete="off" value="Name..." />
    </li>
    <li>
      <label for="Email_From">Email <em>*</em></label><input id="Email_From"  name="_2_Email_From" class="required email replace_text" type="text" autocomplete="off" value="Email..." />
    </li>
    <li>
      <label for="Phone">Phone <em>*</em></label><input id="Phone"  name="_3_Phone" type="text" class="required replace_text" autocomplete="off" />
    </li>
    <li id="textarea">
      <label for="Comments">Comments <em>*</em></label><textarea id="Comments" name="_4_Comments" class="required" rows="" cols="" autocomplete="off"></textarea>
    </li>
    <li class="button">
      <label>&nbsp;</label><input name="Submit" class="submit" value="Submit" type="submit" autocomplete="off" />
    </li>
  </ul>
</form>