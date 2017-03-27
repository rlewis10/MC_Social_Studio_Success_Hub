<script runat="server" language="javascript">
 Platform.Load("Core","1");

   var status = false;
   var id = GUID();
    var email = Request.GetFormField('Email'); //Check to see if the submitter gave us a valid email address
    var firstname = Request.GetFormField('FirstName');
    var lastname = Request.GetFormField('LastName');
    var company = Request.GetFormField('Company');
    var questiontype1 = Request.GetFormField('QuestionType1');
    var questiontype2 = Request.GetFormField('QuestionType2');
    var question = Request.GetFormField('Question');
    var elearning = Request.GetFormField('eLearning');
    var timestamp = Date.now();
    var fpayload = {
                SubscriberKey: id,
                   EmailAddress: email,
                   FirstName: firstname,
                   LastName: lastname,
                   Company: company,
                   QuestionType1: questiontype1,
                   QuestionType2: questiontype2,
                   Question: question,
                   eLearning: elearning,
                   Timestamp: timestamp
          };
  var jsonoutput = Stringify(fpayload);

//The Request.GetFormField() command gets the value from the form POST of the form field with the specified name
//add to DE
 var requestsDE = DataExtension.Init("social_studio_form_data");
   requestsDE.Rows.Add(fpayload);

//add to subscriber list
var subscriberList = List.Init("All Subscribers - 3315");
var status2 = subscriberList.Subscribers.Add(email,{SubscriberKey: id});

//send to google
var url = 'https://docs.google.com/forms/d/1bjh4SU_lHL3i60zUcV5xjY6-FVReFwV-IGC8v49TX_M/formResponse';
var gpayload = "entry.60331601"+"="+id+"&"+
         "entry.1450949137"+"="+firstname+"&"+
               "entry.1771239196"+"="+lastname+"&"+
               "entry.1780634721"+"="+company+"&"+
               "entry.1242786376"+"="+email+"&"+
               "entry.292261482"+"="+questiontype1+"&"+
               "entry.2019899543"+"="+questiontype2+"&"+
               "entry.1317536726"+"="+question+"&"+
               "entry.1195562804"+"="+elearning;

var contentType = 'application/x-www-form-urlencoded; charset=utf-8';
var result = HTTP.Post(url,contentType,gpayload); 
var status = true;
  
// send autoresponder email
var triggeredSend = TriggeredSend.Init("social_success_hub_form_autorespond");
var status1 = triggeredSend.Send(email);  

  
  if (result.StatusCode != 200)
  {status = false;} //Bad response
  else 
  {status = true;} //Good response

</script>
{"Status":"<ctrl:var name=status2 />","Output":[<ctrl:var name=jsonoutput/>]}