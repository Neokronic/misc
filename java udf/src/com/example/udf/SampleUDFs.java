package com.example.udf;

/**
 * Sample StarRocks Java UDFs for BYOC S3-hosted JAR testing.
 *
 * StarRocks scalar UDF contract:
 *   - Public class with a public evaluate() method
 *   - evaluate() can accept any number of supported types:
 *       String, Integer, Long, Double, Boolean, etc.
 *   - Return type is inferred from evaluate()'s return type
 *   - Null inputs: handle explicitly (StarRocks may pass nulls)
 *   - No constructor args needed; StarRocks instantiates via default constructor
 */

// ─────────────────────────────────────────────────
// UDF 1: Reverse a string
//   SQL: CREATE FUNCTION reverse_str(VARCHAR) RETURNS VARCHAR
// ─────────────────────────────────────────────────
class ReverseString {
    public String evaluate(String input) {
        if (input == null) return null;
        return new StringBuilder(input).reverse().toString();
    }
}

// ─────────────────────────────────────────────────
// UDF 2: Mask an email address  →  j***@example.com
//   SQL: CREATE FUNCTION mask_email(VARCHAR) RETURNS VARCHAR
// ─────────────────────────────────────────────────
class MaskEmail {
    public String evaluate(String email) {
        if (email == null || !email.contains("@")) return email;
        int atIdx = email.indexOf('@');
        if (atIdx <= 1) return email;           // too short to mask meaningfully
        char first = email.charAt(0);
        String domain = email.substring(atIdx); // includes the '@'
        return first + "***" + domain;
    }
}

// ─────────────────────────────────────────────────
// UDF 3: Safely divide two doubles (avoids divide-by-zero)
//   SQL: CREATE FUNCTION safe_divide(DOUBLE, DOUBLE) RETURNS DOUBLE
// ─────────────────────────────────────────────────
class SafeDivide {
    public Double evaluate(Double numerator, Double denominator) {
        if (numerator == null || denominator == null) return null;
        if (denominator == 0.0) return null;
        return numerator / denominator;
    }
}
