namespace StatementFile.Domain.Entities
{
    /// <summary>
    /// Represents a bank/branch configuration entry.
    /// Corresponds to the branch record used for product selection and routing.
    /// </summary>
    public class Bank
    {
        public int    BranchCode    { get; private set; }
        public string BranchName    { get; private set; }
        public string BranchPart    { get; private set; }
        public string Ident         { get; private set; }

        private Bank() { }

        public static Bank Create(int branchCode, string branchName, string branchPart, string ident)
        {
            return new Bank
            {
                BranchCode = branchCode,
                BranchName = branchName,
                BranchPart = branchPart,
                Ident      = ident,
            };
        }
    }
}
