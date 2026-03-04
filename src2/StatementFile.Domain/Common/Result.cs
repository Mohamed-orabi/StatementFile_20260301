using System;

namespace StatementFile.Domain.Common
{
    /// <summary>
    /// A discriminated-union result type that carries either a success value
    /// or a structured error, eliminating the need for exception-driven control flow
    /// across use-case boundaries.
    /// </summary>
    public class Result
    {
        public bool    IsSuccess { get; }
        public bool    IsFailure => !IsSuccess;
        public string  Error     { get; }

        protected Result(bool isSuccess, string error)
        {
            if (isSuccess  && error != null)
                throw new InvalidOperationException("A successful result cannot carry an error.");
            if (!isSuccess && string.IsNullOrWhiteSpace(error))
                throw new InvalidOperationException("A failed result must carry an error message.");

            IsSuccess = isSuccess;
            Error     = error;
        }

        public static Result    Ok()              => new Result(true, null);
        public static Result    Fail(string error) => new Result(false, error);
        public static Result<T> Ok<T>(T value)    => new Result<T>(value, true, null);
        public static Result<T> Fail<T>(string error) => new Result<T>(default, false, error);
    }

    /// <summary>
    /// Generic variant that also carries a typed payload on success.
    /// </summary>
    public sealed class Result<T> : Result
    {
        public T Value { get; }

        internal Result(T value, bool isSuccess, string error)
            : base(isSuccess, error)
        {
            Value = value;
        }
    }
}
