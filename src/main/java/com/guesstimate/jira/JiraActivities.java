package com.guesstimate.jira;


import com.google.gson.JsonParser;
import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.ws.commons.util.Base64;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.*;
import java.nio.charset.StandardCharsets;

public class JiraActivities {


    public String createBug(String host, String authToken, String body) throws IOException {

        URL url = new URL("https://" + host + "/rest/api/2/issue/");
        URLConnection urlConnection = url.openConnection();
        urlConnection.setDoInput(true);

        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setDoOutput(true);
        conn.setDoInput(true);

        conn.setRequestMethod("POST");
        conn.setRequestProperty("Authorization", "Basic " + authToken);
        conn.setRequestProperty("Content-Type", "application/json");
        conn.getOutputStream().write(body.getBytes());

        InputStream inputStream = conn.getInputStream();
        return new JsonParser().parse(new InputStreamReader(inputStream)).getAsJsonObject().get("id").getAsString();


    }

    public String getStatus(String host, String authToken, String issueId) throws IOException {

        URL url = new URL("https://" + host + "/rest/api/2/issue/" + issueId);
        URLConnection urlConnection = url.openConnection();
        urlConnection.setDoInput(true);

        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setDoOutput(true);
        conn.setDoInput(true);

        conn.setRequestMethod("GET");
        conn.setRequestProperty("Authorization", "Basic " + authToken);
        conn.setRequestProperty("Content-Type", "application/json");

        InputStream inputStream = conn.getInputStream();
        return new JsonParser().parse(new InputStreamReader(inputStream)).getAsJsonObject().getAsJsonObject("fields").getAsJsonObject("status").get("name").getAsString();


    }

    public String getKey(String host, String authToken, String issueId) throws IOException {

        URL url = new URL("https://" + host + "/rest/api/2/issue/" + issueId);
        URLConnection urlConnection = url.openConnection();
        urlConnection.setDoInput(true);

        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setDoOutput(true);
        conn.setDoInput(true);

        conn.setRequestMethod("GET");
        conn.setRequestProperty("Authorization", "Basic " + authToken);
        conn.setRequestProperty("Content-Type", "application/json");

        InputStream inputStream = conn.getInputStream();
        return new JsonParser().parse(new InputStreamReader(inputStream)).getAsJsonObject().get("key").getAsString();


    }

    public boolean reopenBug(String host, String authToken, String issueId, String reopenCode) throws IOException {

        URL url = new URL("https://" + host + "/rest/api/2/issue/" + issueId + "/transitions?expand=transitions.fields");
        URLConnection urlConnection = url.openConnection();
        urlConnection.setDoInput(true);

        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setDoOutput(true);
        conn.setDoInput(true);

        String body = "{\"transition\": {\"id\": \"" + reopenCode + "\"}}";

        conn.setRequestMethod("POST");
        conn.setRequestProperty("Authorization", "Basic " + authToken);
        conn.setRequestProperty("Content-Type", "application/json");
        conn.getOutputStream().write(body.getBytes());

        return (conn.getResponseCode() == 204);

    }

    public boolean addComments(String host, String authToken, String issueId, String comment) throws IOException {

        URL url = new URL("https://" + host + "/rest/api/2/issue/" + issueId + "/comment");
        URLConnection urlConnection = url.openConnection();
        urlConnection.setDoInput(true);

        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setDoOutput(true);
        conn.setDoInput(true);

        String body = "{\"body\":\"" + comment + "\"}";

        conn.setRequestMethod("POST");
        conn.setRequestProperty("Authorization", "Basic " + authToken);
        conn.setRequestProperty("Content-Type", "application/json");
        conn.getOutputStream().write(body.getBytes());

        return (conn.getResponseCode() == 201);

    }

    public boolean addAttachmentToIssue(String host, String authToken, String issueId, File attachmentFile) throws IOException {

        CloseableHttpClient httpclient = HttpClients.createDefault();

        HttpPost
                httppost = new HttpPost("https://" + host + "/rest/api/latest/issue/" + issueId + "/attachments");
        httppost.setHeader("X-Atlassian-Token", "nocheck");
        httppost.setHeader("Authorization", "Basic " + authToken);

        FileBody fileBody = new FileBody(attachmentFile);

        HttpEntity entity = MultipartEntityBuilder.create()
                .addPart("file", fileBody)
                .build();

        httppost.setEntity(entity);

        CloseableHttpResponse response = httpclient.execute(httppost);

        if (response.getStatusLine().getStatusCode() == 200)
            return true;
        else
            return false;

    }

    private void generateBearerToken(String username, String password) {

        String encoded = Base64.encode((username + ":" + password).getBytes(StandardCharsets.UTF_8));
        System.out.println(encoded);
    }

    public static void main(String[] arg) {

        JiraActivities jiraActivities = new JiraActivities();
        jiraActivities.generateBearerToken(arg[0], arg[1]);
    }

}
