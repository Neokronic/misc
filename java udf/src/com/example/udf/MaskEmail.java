package com.example.udf;

/**
 * UDF: Mask an email address  →  j***@example.com
 *   SQL: CREATE FUNCTION mask_email(VARCHAR) RETURNS VARCHAR
 */
public class MaskEmail {
    public String evaluate(String email) {
        if (email == null || !email.contains("@")) return email;
        int atIdx = email.indexOf('@');
        if (atIdx <= 1) return email;           // too short to mask meaningfully
        char first = email.charAt(0);
        String domain = email.substring(atIdx); // includes the '@'
        return first + "***" + domain;
    }
}
