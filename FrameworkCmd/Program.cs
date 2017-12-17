using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.VisualBasic;
using RefactoringEssentials.CSharp.Converter;

namespace FrameworkCmd
{
    class Program
    {
	    private readonly string toBeReplacedInFilePath;
	    private readonly string replacementInFilePath;

	    private static readonly Dictionary<string, string> DefaultProperties = new Dictionary<string, string>
        {
            {"BuildingInsideVisualStudio", "true"},
            {"SemanticAnalysisOnly", "true"},
        };

	    private Program(string toBeReplacedInFilePath, string replacementInFilePath)
	    {
		    this.toBeReplacedInFilePath = toBeReplacedInFilePath;
		    this.replacementInFilePath = replacementInFilePath;
	    }

	    static void Main(string[] args)
		{
			HackToMakeVs15Work(GetArgOrDefault(args, 3, @"C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise"));

			new Program(GetArgOrDefault(args, 1, "nothing"), GetArgOrDefault(args, 2, "nothing"))
				.WriteConvertedCode(args[0], CancellationToken.None).GetAwaiter().GetResult();
        }

	    private static string GetArgOrDefault(string[] args, int index, string @default)
	    {
		    return args.Length > index ? args[index] : @default;
	    }


	    private static void HackToMakeVs15Work(string vsInstallDir)
	    {
		    Environment.SetEnvironmentVariable("VSINSTALLDIR", vsInstallDir);
		    Environment.SetEnvironmentVariable("VisualStudioVersion", @"15.0");
	    }

	    public async Task WriteConvertedCode(string solutionFilePath, CancellationToken cancellationToken)
        {
	        using (var msWorkspace = MSBuildWorkspace.Create(DefaultProperties))
            {
	            foreach (var message in msWorkspace.Diagnostics.Select(x => x.Message))
                {
                    Console.WriteLine("WARNING: May degrade conversion:\r\n  " + message);
                }
	            var projects = (await GetProjects(solutionFilePath, cancellationToken, msWorkspace)).Where(p => p.Language == LanguageNames.VisualBasic).ToList();
	            if (!projects.SelectMany(p => p.Documents).Any())
	            {
		            throw new Exception("No documents found. This could be due to the warnings above.");
	            }
	            foreach (var project in projects)
	            {
		            var compilation = (VisualBasicCompilation) await project.GetCompilationAsync(cancellationToken);
		            var syntaxTrees = project.Documents.Select(d => d.GetSyntaxTreeAsync(cancellationToken).GetAwaiter().GetResult()).OfType<VisualBasicSyntaxTree>();
		            WriteConversionResult(VisualBasicConverter.ConvertMultiple(compilation, syntaxTrees));
	            }
            }
        }

	    private void WriteConversionResult(Dictionary<string, CSharpSyntaxNode> pathSyntaxNodes)
	    {
		    foreach (var pathSyntaxNode in pathSyntaxNodes)
		    {
				try
			    {
				    var filePathWithReplacement = pathSyntaxNode.Key.Replace(toBeReplacedInFilePath, replacementInFilePath);
				    var newFilePath = Path.ChangeExtension(filePathWithReplacement, "cs");
				    var convertedCode = pathSyntaxNode.Value.NormalizeWhitespace().ToFullString();
				    Directory.CreateDirectory(Path.GetDirectoryName(newFilePath));
				    File.WriteAllText(newFilePath, convertedCode);
			    }
			    catch (Exception e)
			    {
				    Console.WriteLine($"{pathSyntaxNode.Key}\r\n {e}");
			    }
		    }
	    }

	    private static async Task<IEnumerable<Project>> GetProjects(string solutionFilePath, CancellationToken cancellationToken,
            MSBuildWorkspace msWorkspace)
        {
            if (solutionFilePath.EndsWith(".sln"))
            {
                var solution = await msWorkspace.OpenSolutionAsync(solutionFilePath, cancellationToken);
                return solution.Projects;
            }
            return new[] { await msWorkspace.OpenProjectAsync(solutionFilePath, cancellationToken) };
        }


    }
}
