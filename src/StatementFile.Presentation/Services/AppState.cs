namespace StatementFile.Presentation.Services
{
    /// <summary>
    /// Scoped per-browser-session state.
    /// Replaces the WinForms global application state that was implicit in
    /// the single-process desktop model.
    /// </summary>
    public sealed class AppState
    {
        public bool   IsLoggedIn { get; set; }
        public string UserName   { get; set; } = string.Empty;
    }
}
