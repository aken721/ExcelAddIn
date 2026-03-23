using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelAddIn.Skills
{
    public interface ISkill
    {
        string Name { get; }
        string Description { get; }
        List<SkillTool> GetTools();
        Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments);
    }

    public class SkillTool
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
        public List<string> RequiredParameters { get; set; } = new List<string>();
    }

    public class SkillResult
    {
        public bool Success { get; set; }
        public string Content { get; set; }
        public string Error { get; set; }
    }
}