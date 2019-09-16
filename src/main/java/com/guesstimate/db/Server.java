package com.guesstimate.db;

public enum Server {

    QA("", "", "");

    private String host;
    // here value stands for bucket in couchbase, endpoint for api, db for sql
    private String value;
    private String credentials;

    private Server(String host, String value, String credentials) {

        this.host = host;
        this.value = value;
        this.credentials = credentials;
    }

    public String getPassword() {
        return credentials.split(":")[1];
    }

    public String getUserName() {
        return credentials.split(":")[0];
    }

    public String getHost() {
        return host;
    }

    public String getValue() {
        return value;
    }

    public String getCredentials() {
        return credentials;
    }

}
