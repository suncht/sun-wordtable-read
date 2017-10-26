package com.suncht.wordread.exceptions;

public class ParseException extends Exception {

	private static final long serialVersionUID = 939204100093323412L;

	public ParseException(Throwable e) {
		super(e.getMessage(), e);
	}

	public ParseException(String message) {
		super(message);
	}

	public ParseException(String messageTemplate, Object... params) {
		super(String.format(messageTemplate, params));
	}

	public ParseException(String message, Throwable throwable) {
		super(message, throwable);
	}

	public ParseException(Throwable throwable, String messageTemplate, Object... params) {
		super(String.format(messageTemplate, params), throwable);
	}
}
