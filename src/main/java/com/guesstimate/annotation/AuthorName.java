package com.guesstimate.annotation;

public enum AuthorName {

	Nitin("Nitin Simgekar");

	private String authorName;

	AuthorName(String authorName) {
		this.authorName = authorName;
	}

	public String toString() {
		return this.authorName;
	}
}
