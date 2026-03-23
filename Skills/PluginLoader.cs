using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ExcelAddIn.Skills
{
    public class PluginLoader
    {
        public static List<ISkill> LoadSkillsFromAssembly(string assemblyPath, ExcelMcp excelMcp = null)
        {
            var skills = new List<ISkill>();
            try
            {
                var assembly = Assembly.LoadFrom(assemblyPath);
                var skillTypes = Array.FindAll(assembly.GetTypes(), t => typeof(ISkill).IsAssignableFrom(t) && !t.IsAbstract);

                foreach (var type in skillTypes)
                {
                    try
                    {
                        // 尝试创建实例（支持构造函数注入ExcelMcp）
                        var constructor = type.GetConstructor(new[] { typeof(ExcelMcp) });
                        if (constructor != null && excelMcp != null)
                        {
                            var skill = (ISkill)constructor.Invoke(new object[] { excelMcp });
                            skills.Add(skill);
                        }
                        else if (type.GetConstructor(Type.EmptyTypes) != null)
                        {
                            var skill = (ISkill)Activator.CreateInstance(type);
                            skills.Add(skill);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"创建技能实例失败: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载插件失败: {ex.Message}");
            }
            return skills;
        }

        public static List<ISkill> LoadSkillsFromDirectory(string directoryPath, ExcelMcp excelMcp = null)
        {
            var skills = new List<ISkill>();
            if (Directory.Exists(directoryPath))
            {
                foreach (var dll in Directory.GetFiles(directoryPath, "*.dll"))
                {
                    var assemblySkills = LoadSkillsFromAssembly(dll, excelMcp);
                    skills.AddRange(assemblySkills);
                }
            }
            return skills;
        }
    }
}