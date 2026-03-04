namespace StatementFile.Application.DTOs
{
    public sealed class GenerationResultDto
    {
        public bool   Success           { get; set; }
        public string Error             { get; set; }
        public int    FilesGenerated    { get; set; }
        public int    EmailsSent        { get; set; }
        public int    StatementsCount   { get; set; }
        public string OutputDirectory   { get; set; }
    }
}
