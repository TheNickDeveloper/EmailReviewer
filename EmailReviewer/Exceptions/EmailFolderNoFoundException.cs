using System;

namespace EmailReviewer.Exceptions
{
    public class EmailFolderNoFoundException: Exception
    {
        private readonly string _errorMessage;

        public EmailFolderNoFoundException(string errorMessage)
        {
            _errorMessage = errorMessage;
        }

        public override string Message
        {
            get
            {
                return _errorMessage;
            }
        }
    }
}
