package com.guesstimate.utils;

import java.io.IOException;

import org.apache.http.util.TextUtils;

import okhttp3.Credentials;
import okhttp3.Interceptor;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;
import retrofit2.Retrofit;
import retrofit2.converter.gson.GsonConverterFactory;

public class ServiceGenerator extends GuesstimateBase{

	private static OkHttpClient.Builder httpClient = new OkHttpClient.Builder();

	private static Retrofit.Builder builder = new Retrofit.Builder().baseUrl(jenkinsBaseUrl)
			.addConverterFactory(GsonConverterFactory.create());

	private static Retrofit retrofit = builder.build();

	public static <S> S createService(Class<S> serviceClass, String username, String password) {
		String authToken = Credentials.basic(username, password);
		return createService(serviceClass, authToken);

	}

	public static <S> S createService(Class<S> serviceClass, final String authToken) {
		if (!TextUtils.isEmpty(authToken)) {
			AuthenticationInterceptor interceptor = new AuthenticationInterceptor(authToken);

			if (!httpClient.interceptors().contains(interceptor)) {
				httpClient.addInterceptor(interceptor);

				builder.client(httpClient.build());
				retrofit = builder.build();
			}
		}

		return retrofit.create(serviceClass);
	}

	public static class AuthenticationInterceptor implements Interceptor {

		private String authToken;

		public AuthenticationInterceptor(String token) {
			this.authToken = token;
		}

		@Override
		public Response intercept(Chain chain) throws IOException {
			Request original = chain.request();

			Request.Builder builder = original.newBuilder().header("Authorization", authToken);

			Request request = builder.build();
			return chain.proceed(request);
		}
	}
}
