using System.Runtime.Serialization;

namespace MajConsoSuiviSprint.Cli.Helper
{
    [Serializable]
    internal class ExceptFileOpenException : Exception
    {
        public ExceptFileOpenException()
        {
        }

        public ExceptFileOpenException(string? message) : base(message)
        {
        }

        public ExceptFileOpenException(string? message, Exception? innerException) : base(message, innerException)
        {
        }

        protected ExceptFileOpenException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}