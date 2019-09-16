package com.guesstimate.db;

import java.util.concurrent.TimeUnit;

import com.couchbase.client.java.Bucket;
import com.couchbase.client.java.Cluster;
import com.couchbase.client.java.CouchbaseCluster;
import com.couchbase.client.java.query.N1qlQuery;
import com.couchbase.client.java.query.N1qlQueryResult;

public class CouchBase {

	private static Bucket bucket;

	public CouchBase connect(String server) {


		Cluster cluster = CouchbaseCluster.create(Server.valueOf(server).getHost());
		if (Server.valueOf(server).getCredentials().isEmpty())
			bucket = cluster.openBucket(Server.valueOf(server).getValue(), 1, TimeUnit.HOURS);
		else
			bucket = cluster.authenticate(Server.valueOf(server).getUserName(), Server.valueOf(server).getPassword())
					.openBucket(Server.valueOf(server).getValue(), 1, TimeUnit.HOURS);

		return this;

	}

	public N1qlQueryResult query(String query) {
		N1qlQueryResult results = bucket.query(N1qlQuery.simple(query));
		return results;
	}

}
