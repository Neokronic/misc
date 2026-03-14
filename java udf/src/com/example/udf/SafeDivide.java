package com.example.udf;

/**
 * UDF: Safely divide two doubles (avoids divide-by-zero)
 *   SQL: CREATE FUNCTION safe_divide(DOUBLE, DOUBLE) RETURNS DOUBLE
 */
public class SafeDivide {
    public Double evaluate(Double numerator, Double denominator) {
        if (numerator == null || denominator == null) return null;
        if (denominator == 0.0) return null;
        return numerator / denominator;
    }
}
