// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory } = require('botbuilder');

var sp = require("@pnp/sp").sp;
var SPFetchClient = require("@pnp/nodejs").SPFetchClient;

var Invoices = require('./object').Invoices;
const poCard = require('./Card1.json');
const invCard = require('./Card2.json');

var invoices = new Invoices();    

sp.setup({
    sp: {
        fetchClientFactory: () => {
            return new SPFetchClient("https://m365x628217.sharepoint.com/sites/TestTeamsMIP", "c4ab9843-65f4-41e9-8a49-c5e04881f0db", "LtjTLjFxgYFZnG6D1YFfGlMzMZzvwI/BuU4DODu1v+I=");
        },
    },
});

class MyBot extends ActivityHandler {
    constructor() {
        
        super();                         

        this.onMessage(async (context, next) => {

            var _text = "Use commands like 'PO Order' or 'Get PO Order' to get PO data. Use 'new invoice' or 'Create new invoice' to create/upload invoice."

            if(invoices.po_orderid==true){                            
                _text = "PO Order doesn't exists"
                var txt = context.activity.text
                await sp.web.lists.getByTitle("PO_LIST").items.get().then((items) => {
                    items.forEach(element => {
                        if(element.Title==txt)
                        {
                            poCard.body[0].text=element.Title; // PO NUMBER
                            poCard.body[1].columns[1].items[0].text=element.VendorName; //PO Vendor Name
                            poCard.body[1].columns[1].items[1].columns[0].items[0].text="Start Date : "+element.StartDate; //Start Date
                            poCard.body[1].columns[1].items[1].columns[1].items[0].text="End Date : "+element.EndDate; //End Date
                            poCard.body[2].facts[0].value="Total Value : "+element.Value+"/-"; //Value
                            poCard.body[2].facts[1].value="Paid Value : "+element.PaidValue+"/-"; //Paid Value
                            poCard.body[2].facts[2].value="Remaining Value : "+element.RemainingValue+"/-"; //Remaining Value      
                            console.log("found");  
                            _text="";                                                                             
                        }        
                    });    
                });
                                
                invoices.po_orderid=false
                if(_text=="")
                await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(poCard)]
                });
            }

            if(invoices.invoiceid==true){
                _text = "PO Order doesn't exists"
                var txt = context.activity.text
                await sp.web.lists.getByTitle("PO_LIST").items.get().then((items) => {
                    items.forEach(element => {
                        if(element.Title==txt)
                        {
                            invCard.body[0].text="New Invoice Entry for Vendor :  "+element.VendorName; // PO NUMBER
                            invCard.body[1].text=element.Title; // PO NUMBER                                                        
                            _text="";                                                                             
                        }        
                    });    
                });                

                invoices.invoiceid=false
                if(_text=="")
                await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(invCard)]
                });
            }

            if (context.activity.value !== undefined)
            if(context.activity.value.type=='Enter'){
                    context.activity.text="nil";
                    console.log()
                    console.log(context.activity.value)

                    sp.web.lists.getByTitle("INV_LIST").items.add({
                        "Title": invCard.body[1].text,
                        "InvoiceNumber": context.activity.value.invoicenumber,
                        "InvoiceDate":  context.activity.value.date,
                        "InvoiceAmmount" : context.activity.value.ammount
                    });
                    
                    

                    var newpaid = parseInt(context.activity.value.ammount,10)

                    sp.web.lists.getByTitle("PO_LIST").items.get().then((items) => {
                        items.forEach(element => {
                            if(element.Title==invCard.body[1].text)
                            {
                                var id = element.Id;            
                                var paidValue = element.PaidValue;
                                var remainingValue =element.RemainingValue;         
                                console.log(id);              
                                if(newpaid<remainingValue)
                                sp.web.lists.getByTitle("PO_LIST").items.getById(id).update({
                                    PaidValue : (paidValue+newpaid),
                                    RemainingValue : (remainingValue-newpaid)
                                })
                            }        
                        });          
                    })

                    _text="Please attach the invoice copy"
            } 
            
            if(context.activity.attachments){
                _text = "Thank you for attaching invoice"
            }

            if(context.activity.text != undefined)
            if(context.activity.text.toLowerCase().includes('po order')){
                _text = "Enter the PO Order number";
                invoices.po_orderid = true;
            }                        

            if(context.activity.text != undefined)
            if(context.activity.text.toLowerCase().includes('invoice')){
                _text = "Enter the PO number on which invoice to be filled";
                invoices.invoiceid = true;
            }                       


            await context.sendActivity(_text);
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {                    
                    await context.sendActivity("Use commands like 'PO Order' or 'Get PO Order' to get PO data. Use 'new invoice' or 'Create new invoice' to create/upload invoice.");
                }
            }            
            invoices.End();
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.MyBot = MyBot;
