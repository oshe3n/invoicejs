// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory } = require('botbuilder');

var sp = require("@pnp/sp").sp;
var SPFetchClient = require("@pnp/nodejs").SPFetchClient;

var Invoices = require('./object').Invoices;
const poCard = require('./Card1.json');
const invCard = require('./Card2.json');
var invDetail = require('./Card3.json');
var newcard = require('./Card4.json');


var z = '{"type":"ColumnSet","columns":[{"type":"Column","items":[{"type":"TextBlock","text":"date1","wrap":true}],"width":"auto"},{"type":"Column","spacing":"Medium","items":[{"type":"TextBlock","text":"number1","wrap":true}],"width":"stretch"},{"type":"Column","items":[{"type":"TextBlock","text":"am1","wrap":true}],"width":"auto"}]}';

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

            var _text = "Use commands like 'PO' or 'Get PO' to get PO data. Use 'new invoice' or 'Create new invoice' to create/upload invoice. Use 'get invoices' or 'invoices' to list Invoices"
            var flag=true;

            if(invoices.po_orderid==true && invoices.invoiceid==true)
            {
                
                invoices.invoiceid = false;
                invoices.po_orderid = false; 
                
                var txt = context.activity.text;
                
                var json= {};                
                json = Object.assign({},newcard);  
                json.body=[];                
                json.body.push(JSON.parse('{"type":"Container","style":"emphasis","items":[{"type":"ColumnSet","columns":[{"type":"Column","items":[{"type":"TextBlock","size":"Large","weight":"Bolder","text":"**INVOICE SUMMARY**"}],"width":"stretch"}]}],"bleed":true}'));
                json.body.push(JSON.parse('{"type":"Container","items":[{"type":"ColumnSet","columns":[{"type":"Column","items":[{"type":"TextBlock","size":"Large","text":"PO Number : '+txt+'","wrap":true}],"width":"stretch"},{"type":"Column","items":[{"type":"ActionSet","actions":[{"type":"Action.OpenUrl","title":"EXPORT AS PDF","url":"https://m365x628217.sharepoint.com/sites/TestTeamsMIP"}]}],"width":"auto"}]}]}'))
                json.body.push(JSON.parse('{"type":"Container","spacing":"Large","style":"emphasis","items":[{"type":"ColumnSet","columns":[{"type":"Column","items":[{"type":"TextBlock","weight":"Bolder","text":"DATE"}],"width":"auto"},{"type":"Column","spacing":"Large","items":[{"type":"TextBlock","weight":"Bolder","text":"INVOICE NUMBER"}],"width":"stretch"},{"type":"Column","items":[{"type":"TextBlock","weight":"Bolder","text":"AMOUNT"}],"width":"auto"}]}],"bleed":true}'))

                var tot = 0;

                await sp.web.lists.getByTitle("INV_LIST").items.get().then((items) => {
                    items.forEach(element => {                                              
                        if(element.Title==txt)
                        {         
                            var k = JSON.parse(z);                                                                                                           
                            k.columns[0].items[0].text = element.InvoiceDate
                            k.columns[1].items[0].text = element.InvoiceNumber
                            k.columns[2].items[0].text = element.InvoiceAmmount
                            json.body.push(k);                                 
                            tot+= parseInt(element.InvoiceAmmount);
                            /*var json= new Object();
                            json = Object.assign({},invDetail);                                                                          
                            json.body[1].facts[0].value = txt
                            json.body[1].facts[1].value = element.InvoiceNumber                            
                            json.body[1].facts[2].value = element.InvoiceDate
                            json.body[1].facts[3].value = element.InvoiceAmmount                            
                            aatach.push(CardFactory.adaptiveCard(json)) */
                        }        
                    });    
                });                            
                
                json.body.push(JSON.parse('{"type":"Container","style":"emphasis","items":[{"type":"ColumnSet","columns":[{"type":"Column","items":[{"type":"TextBlock","horizontalAlignment":"Right","text":"Total","wrap":true}],"width":"stretch"},{"type":"Column","items":[{"type":"TextBlock","weight":"Bolder","text":"'+tot+'"}],"width":"auto"}]}],"bleed":true}'));

                await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(json)]
                });    

                flag=false;
            }

            if(invoices.po_orderid==true){                            
                _text = "PO doesn't exists"
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
                else
                await context.sendActivity(_text);
                flag=false;
            }

            if(invoices.invoiceid==true){
                _text = "PO doesn't exists"
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
                else
                await context.sendActivity(_text);
                flag=false;
            }

            if (context.activity.value !== undefined)
            if(context.activity.value.type=='Enter'){
                    context.activity.text="nil";
                    console.log()
                    var curr = context.activity.value.date;                    
                    
                    _text="Please attach the invoice copy"                    

                    var newpaid = parseInt(context.activity.value.ammount,10)

                    await sp.web.lists.getByTitle("PO_LIST").items.get().then((items) => {
                        items.forEach(element => {
                            if(element.Title==invCard.body[1].text)
                            {
                                var id = element.Id;            
                                var paidValue = element.PaidValue;
                                var enddate = Date.parse(element.EndDate);
                                var currdate = Date.parse(curr);
                                var remainingValue =element.RemainingValue;         
                                
                                if(newpaid>remainingValue)
                                    _text = "The PO have insufficient balance"
                                
                                if(currdate>enddate)
                                    _text = "Cannot create invoice after expiry, The PO expires "+element.EndDate;

                                if(newpaid<remainingValue && currdate<enddate)
                                {
                                    sp.web.lists.getByTitle("INV_LIST").items.add({
                                        "Title": invCard.body[1].text,
                                        "InvoiceNumber": context.activity.value.invoicenumber,
                                        "InvoiceDate":  context.activity.value.date,
                                        "InvoiceAmmount" : context.activity.value.ammount
                                    });

                                    sp.web.lists.getByTitle("PO_LIST").items.getById(id).update({
                                        PaidValue : (paidValue+newpaid),
                                        RemainingValue : (remainingValue-newpaid)
                                    }) 
                                }                                                                                               
                            }        
                        });          
                    })

                   

                    await context.sendActivity(_text);
                    flag=false;
            } 
            
            if(context.activity.attachments){
                _text = "Thank you for attaching invoice"
                await context.sendActivity(_text);                
                flag=false;
            }

            if(context.activity.text != undefined)
            if(context.activity.text.toLowerCase().includes('po')){
                _text = "Enter the PO Order number";
                invoices.po_orderid = true;
                await context.sendActivity(_text);
                flag=false;
            }                        

            if(context.activity.text != undefined)
            if(context.activity.text.toLowerCase().includes('invoices')){
                _text = "Enter the PO number to list Invoices";
                invoices.invoiceid = true;
                invoices.po_orderid = true;
                context.activity.text='nil'
                await context.sendActivity(_text);
                flag=false;
            }

            if(context.activity.text != undefined)
            if(context.activity.text.toLowerCase().includes('invoice')){
                _text = "Enter the PO number on which invoice to be filled";
                invoices.invoiceid = true;
                await context.sendActivity(_text);
                flag=false;
            }                       


            if(flag)
            {
                await context.sendActivity(_text);                                
            }
            
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {                    
                    
                }
            }            
            invoices.End();
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.MyBot = MyBot;
