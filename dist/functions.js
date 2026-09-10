/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/*!************************************!*\
  !*** ./src/functions/functions.ts ***!
  \************************************/

/// <reference types="office-js" />
/// <reference types="office-runtime" />
/* global CustomFunctions */
/**
 * Returns a spilled row of factorials from 0! up to N!.
 * @customfunction FACTORIALROW
 * @param n The upper bound N. The function returns [0!, 1!, 2!, ..., N!].
 * @returns A 1 x (N+1) row of factorial values.
 */
function factorialRow(n) {
    if (n === null || n === undefined || typeof n !== 'number' || isNaN(n)) {
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, 'N must be a valid number.');
    }
    const upperBound = Math.floor(n);
    if (upperBound < 0) {
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, 'N must be non-negative.');
    }
    // 170! is the largest factorial representable in a double before it overflows to Infinity.
    const MAX_N = 170;
    if (upperBound > MAX_N) {
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, `N must be less than or equal to ${MAX_N}.`);
    }
    const results = [1]; // 0! = 1
    let factorial = 1;
    for (let i = 1; i <= upperBound; i++) {
        factorial *= i;
        results.push(factorial);
    }
    // Excel spills a 2D array: one row containing N+1 columns.
    return [results];
}
CustomFunctions.associate('FACTORIALROW', factorialRow);
CustomFunctions.associate("FACTORIALROW", factorialRow);

/******/ })()
;
//# sourceMappingURL=functions.js.map