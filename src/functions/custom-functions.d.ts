declare namespace CustomFunctions {
  function associate(id: string, functionImpl: (...args: any[]) => any): void;

  enum ErrorCode {
    invalidValue = 'InvalidValue',
    invalidNumberOfArguments = 'InvalidNumberOfArguments',
    invalidRef = 'InvalidReference',
    notAvailable = 'NotAvailable',
    divideByZero = 'DivideByZero',
    invalidName = 'InvalidName',
  }

  class Error {
    constructor(code: CustomFunctions.ErrorCode, message?: string);
    code: CustomFunctions.ErrorCode;
    message?: string;
  }
}
