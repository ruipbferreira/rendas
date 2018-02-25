package model;

public class Fraction {
    private String name;
    private String nif;
    private String address;
    private String value;
    private String addressCode;
    private String fractionCode;
    private String articleCode;
    private String share;

    public Fraction() {
    }

    public Fraction(String name, String nif, String address, String value, String addressCode, String fractionCode, String articleCode, String share) {
        this.name = name;
        this.nif = nif;
        this.address = address;
        this.value = value;
        this.addressCode = addressCode;
        this.fractionCode = fractionCode;
        this.articleCode = articleCode;
        this.share = share;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getNif() {
        return nif;
    }

    public void setNif(String nif) {
        this.nif = nif;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getAddressCode() {
        return addressCode;
    }

    public void setAddressCode(String addressCode) {
        this.addressCode = addressCode;
    }

    public String getFractionCode() {
        return fractionCode;
    }

    public void setFractionCode(String fractionCode) {
        this.fractionCode = fractionCode;
    }

    public String getArticleCode() {
        return articleCode;
    }

    public void setArticleCode(String articleCode) {
        this.articleCode = articleCode;
    }

    public String getShare() {
        return share;
    }

    public void setShare(String share) {
        this.share = share;
    }
}
