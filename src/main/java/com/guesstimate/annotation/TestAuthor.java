package com.guesstimate.annotation;

import static java.lang.annotation.ElementType.METHOD;
import static java.lang.annotation.ElementType.TYPE;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ METHOD, TYPE })
@Retention(RetentionPolicy.RUNTIME)
public @interface TestAuthor {

	AuthorName name();

	String createdDate() default "N/A";

}
