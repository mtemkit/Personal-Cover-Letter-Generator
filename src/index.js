/*Quick Set-Up Reminder:

- Running script: npm run start
- Use nodemon with "yarn add nodemon"/"npm install -g nodemon" to
automatically update the console and refresh the application whenever I press "save"
- To use parcel bundler: npm install parcel-bundler --save-dev 
- key name for ".env" is OPENAI_API_KEY
*/




//-------------------------------------------------------------------------------------------------------------------------------------
//PACKAGES USED!!!


import { Document, Packer } from "docx";
import { saveAs } from "file-saver";
import fs from 'fs'
const d = new Date();
import { Configuration, OpenAIApi } from "openai";
import * as dotenv from 'dotenv'; //hiding api keys or private info into .env files
dotenv.config({path: '__dirname+../.env'})
import { Buffer } from 'buffer';




//-------------------------------------------------------------------------------------------------------------------------------------
//VARIABLES SET EARLY & FUNCTIONS

//"A way that I've innovated is,"
//"Additionally, I like problem solving and finding new solution,""
//"I also like to continue learning about programming and improving my skills such as with".
//just a general guideline. The focus should be on including key words, and keeping the paragraphs professional and concise, with unique and seamless transitions.  

let coverLetter = {};
let jobKeyExpectations = ''.concat(
`Write me a total of 3 paragraphs', Each paragraph must be under 100 words and be professional. There must be a new line at the end of each paragraph.`,
`The paragraphs that are written are meant to be in the body of a cover letter. I only want the three paragraphs mentioned.`,
`For the first paragraph, it should start with the best transition similar to ""A way that I've innovated is,".`,
`The theme of the first paragraph will be about the positive impact made at work through my communicatation skills.`,
`The situation for the first paragraph will be about how I made a postive impact as a WordPress Developer by saving my employer over $200 in site plugins.`,
`For the second paragraph, it should start with the best transition similar to "Additionally, I've developed my problem solving skills by".`,
`The theme of the second paragraph will be about how I have developed problem solving skills through dozens of tickets and tougher programming tasks and projects.`,
`The situation for the second paragraph will be about how I managed to tackle completing a deprecated login page for Internet Explorer users of the company application during 
my co-op as a Software Developer.`,
`For the third paragraph, it should start with the best transition similar to "I've also continue learning about programming improving my skills through".`,
`The theme of the third paragraph will be about how I've continued feeding my desire to learn more about programming and how I've improved my programming skills through 
personal projects.`,
`The situation of the third paragraph will be about how I managed to develop my programming skills through developed and tested a social media website using 
LAMP stack (Linux, Apache, MySQL, and PHP).`,
`For each of the paragraphs from the three paragraphs requested, make sure to incorporate as many key words from the job requirement information below into the three paragraphs 
while still meeting the previously mentioned requirements.`,
`Now, any below this point should be considered as the list of job requirements and should be treated as keywords to be used and added to the three paragraphs 
(without surpassing 100 words per paragraph): `
);

console.log(jobKeyExpectations);

const configuration = new Configuration({
  apiKey: process.env.OPENAI_API_KEY,
});
const openai = new OpenAIApi(configuration);

function wordBeforePosition(position) {
  // Get the first word of the sentence by splitting the sentence on spaces
  // and taking the first element of the resulting array
  const firstWord = position.split(' ')[0];

  // Check if the first word starts with a vowel (a, e, i, o, or u)
  const vowelRegex = /^[aeiou]/i;
  if (vowelRegex.test(firstWord)) {
    return 'an';
  } else {
    return 'a';
  }
}

function wordBeforeCompany(company) {
  // Normalize the input string by converting it to lower case
  // and removing leading and trailing whitespace
  company = company.toLowerCase().trim();

  // Check if the company name starts with "a", "an", or "the"
  if (company.startsWith("a ") || company.startsWith("an ") || company.startsWith("the ")) {
    return "";
  }

  // Check if the company name contains a group-indicating word (any sort of organization)
  const groupWords = ["organization", "agency", "firm", "bureau", "group", "department", "board",
  "branch", "team", "club", "business", "service", "channel", "office", "government", "division",
  "comission", "shop", "store", "desk"];
  for (const word of groupWords) {
    if (company.includes(word)) {
      return "the ";
    }
  }

  // Otherwise, return an empty string
  return "";
}


//function to add a period to companies ending in inc, ltd, or co where needed in the letter
function addPeriod(string) {
  const words = string.split(' '); // split the string into an array of words
  if (words[words.length - 1].toLowerCase() === 'inc' || 
  words[words.length - 1].toLowerCase() === 'co' || words[words.length - 1].toLowerCase() === 'ltd') { 
    // check if the last word is "inc", "co", or "ltd"
    words[words.length - 1] += '.'; // add a period to the last word0
  }
  return words.join(' '); // join the words back into a string and return it
}

//function to remove a period from companies ending in inc., ltd., or co. where needed in the letter
function removePeriod(string) {
  const words = string.split(' '); // split the string into an array of words
  if (words[words.length - 1].toLowerCase() === 'inc.' || words[words.length - 1].toLowerCase() === 'co.' ||
   words[words.length - 1].toLowerCase() === 'ltd.') {
    // check if the last word is "inc.", "co.", or "ltd."
    words[words.length - 1] = words[words.length - 1].slice(0, -1); // remove the period from the last word
  }
  return words.join(' '); // join the words back into a string and return it
}

function formatCommaSeparatedValues(string) {

  if (string.trim() == '' || string.trim() == ','){
    return ""
  }

  // Split the string on commas
  const values = string.split(',');

  // Strip any leading or trailing whitespace from each value
  const strippedValues = values.map(value => value.trim());

  // Join the values with a comma and a space
  return strippedValues.join(', ')+", ";
}

function capitalizeEachWord(companyName){
  return companyName.split(' ').map(s => s.charAt(0).toUpperCase() + s.slice(1)).join(' ');
}

function expressInterest(companyName, companyDescription) {
  let interestPhrases = [
    "I am drawn to",
    "I am interested in",
    "I am excited about",
    "I am impressed by",
    "I am captivated by"
  ];

  let randomInterestPhrase = interestPhrases[Math.floor(Math.random() * interestPhrases.length)];

  return `${randomInterestPhrase} ${wordBeforeCompany(companyName)}${capitalizeEachWord(companyName)}'s reputation for ${companyDescription}`;
}

function processText(text) {
  // Remove leading and trailing spaces
  text = text.trim();

  // Split the text into an array of paragraphs
  let paragraphs = text.split('\n');

  // Remove leading and trailing spaces from each paragraph
  paragraphs = paragraphs.map(p => p.trim());

  // Remove empty strings from the array
  paragraphs = paragraphs.filter(p => p !== '');

  return paragraphs;
}

//function for saving the document to a file
function saveDocumentToFile(doc, fileName) {
  // Create new instance of Packer for the docx module
  const packer = new Packer();
  // Create a mime type that will associate the new file with Microsoft Word
  const mimeType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  // Create a Blob containing the Document instance and the mimeType
  packer.toBlob(doc).then(blob => {
    const docblob = blob.slice(0, blob.size, mimeType);

    // Save the file using saveAs from the file-saver package
    saveAs(docblob, fileName);
  });
}



 




//-------------------------------------------------------------------------------------------------------------------------------------
//CLICK EVENT FOR GENERATING A DOCUMENT (TWO BUTTONS)!!!
$('#generate-sample')[0].addEventListener(
  "click",
  (event) => {
    // Prevent the form from reloading the page on submit
    event.preventDefault();

    let positionName = "Web Developer";
    let companyName = "ABC Organization";
    let companyDescription = "service excellence and technology innovation in the global air transportation industry"; //what I like about the company
    //let bodyParagraphs = ""; //three body paragraphs
    let jobKeyRequirements = "Bilingual Imperative: BBB/BBB o	Reliability/Security: Reliability, Secret o	Willingness to work overtime o	Willingness to travel o	Willingness to work on evenings and/or on weekends o	Willingness to work shifts or a flexible schedule (including weekends and statutory holidays)"; 
    let jobKeyLanguages = "java,c++"; //key languages of the position
    let prompt = `${jobKeyExpectations} ###${jobKeyRequirements}###`;

    // Store the values in the cover letter object for later use
    coverLetter = {
      positionName: positionName,
      companyName: companyName,
      companyDescription: companyDescription,
      prompt: prompt,
      jobKeyLanguages: jobKeyLanguages,
      bodyParagraphs: []
    }

    openai.createCompletion({
      model: "text-davinci-003",
      prompt: coverLetter.prompt,
      max_tokens: 350,
      top_p: 1.0,
      frequency_penalty: 0.0,
      presence_penalty: 0.0,
    }).then((res) => 
    {
      coverLetter.bodyParagraphs = processText(res.data.choices[0].text);
      console.log(coverLetter.bodyParagraphs)
      generateWordDocument(event);
    });
  },
  false
);  

$('#generate-unique').submit(function(event) {
  // Prevent the form from reloading the page on submit
  event.preventDefault();
  console.log('hey');

  let positionName = $("#position-name").val();
  let companyName = $("#company-name").val();
  let companyDescription = $("#company-description").val(); //what I like about the company
  let jobKeyRequirements = $("#job-key-requirements").val(); //key requirements of the position
  let jobKeyLanguages = $("#job-key-languages").val(); //key languages of the position
  let prompt = `${jobKeyExpectations} ###${jobKeyRequirements}###`;

  // Store the values in the cover letter object for later use
  coverLetter = {
    positionName: positionName,
    companyName: companyName,
    companyDescription: companyDescription,
    jobKeyRequirements: jobKeyRequirements,
    prompt: prompt,
    jobKeyLanguages: jobKeyLanguages,
    bodyParagraphs: []
  };
  
  openai.createCompletion({
    model: "text-davinci-003",
    prompt: coverLetter.prompt,
    max_tokens: 350,
    top_p: 1.0,
    frequency_penalty: 0.0,
    presence_penalty: 0.0,
  }).then((res) => 
  {
    coverLetter.bodyParagraphs = processText(res.data.choices[0].text);
    generateWordDocument(event);
  });

});  










//-------------------------------------------------------------------------------------------------------------------------------------
//FULL FUNCTION FOR A GENERATED WORD DOCUMENT (CALLED DURING BUTTON CLICK EVENT)

//CREATING THE DOC & FONT COLOR AND TYPE
function generateWordDocument(event) {
  // Create a new instance of Document for the docx module
  let doc = new Document();
  doc.theme = {
    font: {
      normal: {
        family: "Calibri",
        color: "000000"
      },
      header: { family: "Calibri" }
    },
    title: {
      color: "000000"
    },
    headings: {
      one: {
        color: "000000"
      },
      two: {
        color: "000000"
      }
    }
  };


  //-------------------------------------------------------------------------------------------------------------------------------------
  /*TEXT STYLING TYPES (EX. CUSTOM NORMAL & CUSTOM NORMAL CENTRE), ALIGNMENT, 
  SPACING, FONT SIZING (DIVIDE SIZE NUMBER BY 2 FOR ACTUAL TRANSLATED FONT SIZE) ETC.*/
  doc.Styles.createParagraphStyle("customHeading1", "Custom Heading 1")
    .basedOn("Heading 1")
    .next("Normal")
    .quickFormat()
    .font(doc.theme.font.header.family)
    .size(32)
    .bold()
    .color(doc.theme.headings.one.color)
    .spacing({ after: 250 });
  doc.Styles.createParagraphStyle("customHeading2", "Custom Heading 2")
    .basedOn("Heading 2")
    .next("Normal")
    .quickFormat()
    .font(doc.theme.font.header.family)
    .size(26)
    .bold()
    .color(doc.theme.headings.two.color)
    .spacing({ after: 150 });
  doc.Styles.createParagraphStyle("customTitle", "Custom Title")
    .basedOn("Title")
    .next("Normal")
    .quickFormat()
    .font(doc.theme.font.header.family)
    .size(56)
    .bold()
    .color(doc.theme.font.normal.color)
    .spacing({ after: 250 });
  doc.Styles.createParagraphStyle("customSubtitle", "Custom Subtitle")
    .basedOn("Subtitle")
    .next("Normal")
    .quickFormat()
    .font(doc.theme.font.header.family)
    .size(22)
    .color(doc.theme.font.normal.color)
    .spacing({ after: 150 });
    doc.Styles.createParagraphStyle("customNormalNoSpacing", "Custom Normal No Spacing For Address at Top")
    .basedOn("Normal")
    .quickFormat()
    .font(doc.theme.font.normal.family)
    .size(22)
    .color(doc.theme.font.normal.color)
    .spacing({ after: 10 });
  doc.Styles.createParagraphStyle("customNormalCenter", "Custom Normal Center for Subject")
    .basedOn("Normal")
    .quickFormat()
    .font(doc.theme.font.normal.family)
    .size(22)
    .color(doc.theme.font.normal.color)
    .spacing({ after: 150 })
    .center()
    .bold()
  doc.Styles.createParagraphStyle("customNormal", "Custom Normal for Body")
    .basedOn("Normal")
    .quickFormat()
    .font(doc.theme.font.normal.family)
    .size(22)
    .color(doc.theme.font.normal.color)
    .spacing({ after: 195 });
  doc.Styles.createParagraphStyle("customNormalExtraSpacing", "Custom Normal Extra Spacing for Final Thanks")
    .basedOn("Normal")
    .quickFormat()
    .font(doc.theme.font.normal.family)
    .size(22)
    .color(doc.theme.font.normal.color)
    .spacing({ before: 75, after: 150 })


//-------------------------------------------------------------------------------------------------------------------------------------
  //THE ACTUAL TEXT IN THE DOC!!! ADD VARIABLES OR SPACES HERE!!! 

  //Address Letter at the top
  doc.createParagraph(`${d.toLocaleString('default', { month: 'long' })} ${d.getDate()}, ${d.getFullYear()}`).style("customNormalNoSpacing"); //add self-updating date variable
  doc.createParagraph("").style("customNormalNoSpacing");
  doc.createParagraph("Hiring Manager").style("customNormalNoSpacing");
  doc.createParagraph(`${capitalizeEachWord(coverLetter.positionName)}`).style("customNormalNoSpacing"); //add variable
  doc.createParagraph(`${addPeriod(capitalizeEachWord(coverLetter.companyName))}`).style("customNormalNoSpacing"); //add variable
  doc.createParagraph("K1V2C7").style("customNormalNoSpacing"); //if I find an address
  doc.createParagraph("Ottawa, Ontario").style("customNormalNoSpacing");

  //Letter Subject
  doc.createParagraph(`Re: ${capitalizeEachWord(coverLetter.positionName)}`).style("customNormalCenter");

  //Letter Body
  doc.createParagraph("Dear Hiring Manager,").style("customNormal"); //addressing the reciever of the letter

  //First Paragraph
  doc
    .createParagraph(
      `Please accept this letter as my application for ${wordBeforePosition(coverLetter.positionName)} ${coverLetter.positionName.toLowerCase()} position with ${wordBeforeCompany(coverLetter.companyName)}${capitalizeEachWord(removePeriod(coverLetter.companyName))}. With over three years of programming experience, I am confident in my ability to excel in this role and make a valuable contribution to your team.`
    )
    .style("customNormal");

  //Second Paragraph
  doc
    .createParagraph(
      `As a graduating IAWD student from Algonquin College with a strong foundation in programming languages including ${capitalizeEachWord(formatCommaSeparatedValues(coverLetter.jobKeyLanguages))}JavaScript, C#, SQL, Python, HTML, and CSS, I have gained experience in software development and computer science through my co-op and team projects. In addition, I am bilingual (fluent in both English and French) and am capable of solving complex problems and writing concise technical summaries in a professional setting. I have also developed skills in building, testing, maintaining, and deploying code from scratch.`
    )
    .style("customNormal");

  //Third Paragraph
  doc
    .createParagraph(
      `${expressInterest(coverLetter.companyName, coverLetter.companyDescription)}, as it aligns with my desire to use my programming skills to make a positive impact and innovate in the world. ${coverLetter.bodyParagraphs[0]}`
      )
    .style("customNormal");

  //Fourth Paragraph
  doc
  .createParagraph(
    `${coverLetter.bodyParagraphs[1]}`
    )
  .style("customNormal");

  //Fifth Paragraph
  doc
  .createParagraph(
    `${coverLetter.bodyParagraphs[2]}`
    )
  .style("customNormal");


  /*THESE PARTS NEED TO BE MODIFIED!!! 
  I NEED TO OUTPUT AN ARRAY WITH EACH PARAGRAPH AS A VALUE WHEN GIVEN THREE PARAGRAPHS AS INPUT.*/

  /*
   //Fourth Paragraph
   doc
   .createParagraph(
     "My strong programming foundation was developed in the IAWD program where I explored Linux OS in-depth and created projects in both frontend and backend languages such as HTML, CSS, JS, C#, PHP, SQL, and Python. By exploring a wide array of principles, I was able to build, maintain, and deploy several projects in class from scratch such as a Course Registration App in C# and personal projects such as a Social Media App with MERN stack (MongoDB, Express, React.JS, and Node.JS). I further refined my skills as a Software Developer at I-Sight through completing dozens of fix and feature tickets which were reviewed and merged into the Node.JS base application."
   )
   .style("customNormal");
   //Fifth Paragraph
   doc
   .createParagraph(
     "I was able to refine my communication skills in the IAWD program where I exercised patience and open-mindedness as a team lead during discussions and collaborated with team members using Zoom, Microsoft Teams, and GitHub. As a WordPress developer serving bilingual content in both English and French, I was challenged to listen actively, be concise, and take initiative. At I-Sight, I learned the value of giving progress updates and getting feedback to improve quality of work through daily Slack meetings."
   )
   .style("customNormal");
    //Sixth Paragraph
    doc
    .createParagraph(
      "Throughout my programming journey, I’ve developed my problem-solving skills and learned to improve them through both experience and a deliberate effort to deal with problems systemically. This allows me to adapt quickly since I have a solid approach to address blockers. My co-op at I-Sight played a big part in improving this since I’ve learned to quickly write out questions stemming from a ticket, discuss with the team expert to understand the change, examine similar app features, and discuss wider scope implications. This includes knowing when to use Google and being willing to ask for help when blocked."
    )
    .style("customNormal");
  */


    //Conclusion
    doc
    .createParagraph(
      "I believe that my experience, qualifications, and drive to improve make me an ideal candidate for this position and I am excited to bring my skills and enthusiasm to your team."
    )
    .style("customNormal");
    //doc.createParagraph("").style("customNormal"); // this is a space

    //Final Thank you
    doc.createParagraph("Thank you for considering my application.").style("customNormalExtraSpacing");
    //doc.createParagraph("").style("customNormal"); // this is a space
    doc.createParagraph("Sincerely,").style("customNormalExtraSpacing");
    //doc.createParagraph("").style("customNormal"); // this is a space
    doc.createImage(fs.readFileSync('./images/signature.jpg'), 118, 50); //factor of 2.35, first value is width, second is height
    //doc.createParagraph("").style("customNormal"); // this is a space
    doc.createParagraph("Mohamed Temkit").style("customNormalExtraSpacing");


  // Call saveDocumentToFile with the document instance and a filename
  saveDocumentToFile(doc, "Practice Cover Letter - Mohamed Temkit.docx");
  //saveDocumentToFile(doc, "New Document.pdf");
}
//end of the generate documention function