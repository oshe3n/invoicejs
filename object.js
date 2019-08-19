class Invoices {
    PO_Order(po_orderid) {        
        this.po_orderid = po_orderid;        
    }
    End(){
        this.po_orderid = null;
        this.invoiceid = null;
    }
}

module.exports.Invoices = Invoices;