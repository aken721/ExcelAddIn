using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelAddIn.Skills
{
    public class SkillManager
    {
        private List<ISkill> _skills = new List<ISkill>();
        private Dictionary<string, ISkill> _toolToSkillMap = new Dictionary<string, ISkill>();

        public void LoadSkill(ISkill skill)
        {
            _skills.Add(skill);
            
            // 建立工具名称到技能的映射
            foreach (var tool in skill.GetTools())
            {
                _toolToSkillMap[tool.Name] = skill;
            }
        }

        public List<SkillTool> GetAllTools()
        {
            var allTools = new List<SkillTool>();
            foreach (var skill in _skills)
            {
                allTools.AddRange(skill.GetTools());
            }
            return allTools;
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            if (_toolToSkillMap.TryGetValue(toolName, out var skill))
            {
                return await skill.ExecuteToolAsync(toolName, arguments);
            }
            return new SkillResult { Success = false, Error = $"Tool {toolName} not found" };
        }

        public List<ISkill> GetLoadedSkills()
        {
            return _skills;
        }

        public bool IsToolAvailable(string toolName)
        {
            return _toolToSkillMap.ContainsKey(toolName);
        }
    }
}
