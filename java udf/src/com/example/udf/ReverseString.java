package com.example.udf;

/**
 * UDF: Reverse a string
 *   SQL: CREATE FUNCTION reverse_str(VARCHAR) RETURNS VARCHAR
 */
public class ReverseString {
    public String evaluate(String input) {
        if (input == null) return null;
        return new StringBuilder(input).reverse().toString();
    }
}
