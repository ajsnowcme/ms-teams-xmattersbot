

/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");
//var http = require('http');
var request = require('request');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// Create your bot with a function to receive messages from the user
//var bot = new builder.UniversalBot(connector);


var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);

bot.dialog('/', [
    function (session) {
        builder.Prompts.text(session, "Not a triggered word.  Try Help.");
        session.endDialog();
    },
    
        
]);

bot.dialog('Hello', [
    function (session) {
        builder.Prompts.text(session, "Hello... What's your name?");
    },
    function (session, results) {
        session.userData.name = results.response;
        builder.Prompts.number(session, "Hi " + results.response + ", How many years have you been coding?"); 
    },
    function (session, results) {
        session.userData.coding = results.response;
        builder.Prompts.choice(session, "What language do you code Node using?", ["JavaScript", "CoffeeScript", "TypeScript"]);
    },
    function (session, results) {
        session.userData.language = results.response.entity;
        session.send("Got it... " + session.userData.name + 
                    " you've been programming for " + session.userData.coding + 
                    " years and use " + session.userData.language + ".");
        session.endDialog();
    }
]).triggerAction({ matches: /^Hello$/i });

bot.dialog('help', [
    function (session) {
       var text = session.message.text;
       var command = text.split(" ")[0];
       var extras = text.split(command+" ")[1];
       //builder.Prompts.text(session, "Hello... What's your name?");
       help(extras,session);
    }
    
]).triggerAction({ matches: /(Help)\s(.*).*/i });



bot.dialog('engageButtonClick', [
        function (session, args, next) {

            var utterance = args.intent.matched[0];
            var engageMethod = /(SMS|E-Mail|Any Method)/i.exec(utterance);
            var engageType = /\b(Critical Incident|Invite to chat)\b/i.exec(utterance);
            var recipientType = /\b(Directly)\b/i.exec(utterance);
            var contactType = session.dialogData.contactType = {
                utterance: utterance,
                endpoint: "engage",
                engageMethod: engageMethod ? engageMethod[0].toLowerCase() : null,
                engageType: engageType ? engageType[0].toLowerCase() : null,
                target: utterance.split(" ")[1] ? utterance.split(" ")[1] : null,
                recipientType: recipientType ? recipientType[0].toLowerCase()+" " : "",
            };

            //TODO: ensure group exists

            if(contactType.engageType){
                next();
            }else{
                var msg = new builder.Message(session);
                msg.attachments([
                    new builder.HeroCard(session)
                        .title("Engagement Type")
                        .subtitle("Choose the type of engagement")
                        .buttons([
                            builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident", "Critical Incident"),
                            builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Invite to chat", "Invite to chat")
                        ])
                ]);
                session.send(msg).endDialog();
            } 
        },
        function (session, args, next) {
            var contactType = session.dialogData.contactType;
            var utterance = contactType.utterance;

            if(contactType.engageType == "critical incident"){

                var engagePriority = /(High|Medium|Low)/i.exec(utterance);
                contactType.engagePriority = engagePriority ? engagePriority[0].toLowerCase() : null
                session.dialogData.contactType = contactType;

                if(contactType.engagePriority){
                    next();
                }else{
                    var msg = new builder.Message(session);
                    msg.attachments([
                        new builder.HeroCard(session)
                            .title("Incident Priority")
                            .subtitle("Choose the priority of incident")
                            .buttons([
                                builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident with High Priority", "High"),
                                builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident with Medium Priority", "Medium"),
                                builder.CardAction.imBack(session, "Engage "+contactType.target+" "+contactType.recipientType+"Critical Incident with Low Priority", "Low")
                            ])
                    ]);
                    session.send(msg).endDialog();
                } 
            }else{
                next();
            }
        },
        function (session, results) {
            var contactType = session.dialogData.contactType;
            contactType.recipientType = contactType.recipientType.trim();

                postData(contactType.target, session);
                //engage(contactType.target,session,direct);
        }
    ]).triggerAction({ matches: /(Engage)\s(.*).*/i });


function postData(contact, session){
    request.post(
//       'https://advisors.na5.xmatters.com/api/integration/1/functions/0fbe9db0-5e22-4661-af53-c81b88528583/triggers?apiKey=9d4c74a4-a844-4dbd-a3ee-3d15ffb9a499',
       'https://olin.cs1.xmatters.com/api/integration/1/functions/a49b2eb3-1219-4a9b-bf81-b0cd5bae6e18/triggers?apiKey=d2165aba-66fb-4ba7-8c67-be59460d04b5',
       { json: { recipients: contact, session: session } },
       function (error, response, body) {
           if (!error && response.statusCode == 200) {
               console.log(body)
           }
       }
   );  
}

function help(targets,session){
        var helpText = "**You can do the following commands:**\n\n";
        helpText += ". \n\n";
        helpText += "**help:** Displays this help\n\n";
        helpText += "**oncall [group]:** Displays who's on call\n\n";
        helpText += "**engage [group]:** Invite people to the chat\n\n";
        helpText += "**confCall:** Creates a conference bridge\n\n";

        postToChannel(session,helpText,"markdown");
    }

function postToChannel(session, text,type){
        var msg = new builder.Message(session);
        msg.text(text);
        if(!!type){
            console.log(type);
            msg.textFormat(type);
        }
        msg.textLocale('en-US');
        console.log(msg);
        bot.send(msg);
    }









        
       
