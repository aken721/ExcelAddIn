using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Skills
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
        public List<string> Suggestions { get; set; } = new List<string>();
        public bool RequiresUserDecision { get; set; }

        public static SkillResult FromError(string error, List<string> suggestions = null, bool requiresUserDecision = false)
        {
            return new SkillResult
            {
                Success = false,
                Error = error,
                Suggestions = suggestions ?? new List<string>(),
                RequiresUserDecision = requiresUserDecision
            };
        }
    }
}