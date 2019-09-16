package com.guesstimate.utils;

import okhttp3.ResponseBody;
import retrofit2.Call;
import retrofit2.http.GET;
import retrofit2.http.Url;

public interface RestServices {

	@GET
	public Call<ResponseBody> getIpa(@Url String url);

}
