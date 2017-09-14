package com.spicejet.threshold;

import java.util.Date;

public class Book {

    private String docNo;
    private String docType;
    private String revision;
    private String issuedBy;
    private String effTitle;
    private String type;
    private String base;
    private String dim;
    private int intAmount;
    private Date dateAmount;


    public static Book getCopiedBookInstance(Book book) {
        Book value = new Book();
        value.setDocNo(book.getDocNo());
        value.setDocType(book.getDocType());
        value.setRevision(book.getRevision());
        value.setIssuedBy(book.getIssuedBy());
        value.setEffTitle(book.getEffTitle());
        value.setType(book.getType());
        value.setBase(book.getBase());
        value.setDim(book.getDim());
        value.setIntAmount(book.getIntAmount());
        value.setDateAmount(book.getDateAmount());
        return value;
    }

    public String getDocNo() {
        return docNo;
    }

    public void setDocNo(String docNo) {
        this.docNo = docNo;
    }

    public String getDocType() {
        return docType;
    }

    public void setDocType(String docType) {
        this.docType = docType;
    }

    public String getRevision() {
        return revision;
    }

    public void setRevision(String revision) {
        this.revision = revision;
    }

    public String getIssuedBy() {
        return issuedBy;
    }

    public void setIssuedBy(String issuedBy) {
        this.issuedBy = issuedBy;
    }

    public String getEffTitle() {
        return effTitle;
    }

    public void setEffTitle(String effTitle) {
        this.effTitle = effTitle;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getBase() {
        return base;
    }

    public void setBase(String base) {
        this.base = base;
    }

    public String getDim() {
        return dim;
    }

    public void setDim(String dim) {
        this.dim = dim;
    }

    public int getIntAmount() {
        return intAmount;
    }

    public void setIntAmount(int intAmount) {
        this.intAmount = intAmount;
    }

    public Date getDateAmount() {
        return dateAmount;
    }

    public void setDateAmount(Date dateAmount) {
        this.dateAmount = dateAmount;
    }
}
