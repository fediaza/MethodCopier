using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Community.VisualStudio.Toolkit;
using EnvDTE;
using ICSharpCode.Decompiler;
using ICSharpCode.Decompiler.CSharp;
using ICSharpCode.Decompiler.CSharp.Syntax;
using ICSharpCode.Decompiler.Metadata;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Formatting;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.LanguageServices;
using Microsoft.VisualStudio.Text;

using Solution = Microsoft.CodeAnalysis.Solution;

namespace MethodCopier
{
    [Command(PackageIds.MyCommand)]
    internal sealed class CopyMethodCommand : BaseCommand<CopyMethodCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            try
            {
                var docView = await VS.Documents.GetActiveDocumentViewAsync();
                if (docView?.TextView == null) return;

                var selection = docView.TextView.Selection.SelectedSpans.FirstOrDefault();
                if (selection.IsEmpty) return;

                var document = await docView.TextBuffer.GetRoslynDocumentAsync();
                if (document == null) return;

                var solution = document.Project.Solution;
                var semanticModel = await document.GetSemanticModelAsync();
                var root = await document.GetSyntaxRootAsync();
                var textSpan = new TextSpan(selection.Start, selection.Length);
                var node = root.FindNode(textSpan);

                var methodSymbol = await GetMethodSymbolAsync(semanticModel, node);
                if (methodSymbol == null) return;

                var allMethodSymbols = await FindAllMethodDependenciesAsync(methodSymbol, solution);
                var sourceText = await GenerateSourceWithDependenciesAsync(allMethodSymbols, solution);

                Clipboard.SetText(sourceText);
                await VS.StatusBar.ShowMessageAsync($"Copied {allMethodSymbols.Count} methods to clipboard!");
            }
            catch (Exception ex)
            {
                await ex.LogAsync();
                await VS.MessageBox.ShowErrorAsync("Error copying methods", ex.Message);
            }
        }

        private static async Task<IMethodSymbol> GetMethodSymbolAsync(SemanticModel semanticModel, SyntaxNode node)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            // Check method declaration
            if (semanticModel.GetDeclaredSymbol(node) is IMethodSymbol declaredSymbol)
                return declaredSymbol;

            // Check method invocation
            var symbolInfo = semanticModel.GetSymbolInfo(node);
            return symbolInfo.Symbol as IMethodSymbol ?? symbolInfo.CandidateSymbols.FirstOrDefault() as IMethodSymbol;
        }

        private static async Task<HashSet<IMethodSymbol>> FindAllMethodDependenciesAsync(
            IMethodSymbol rootMethod, Solution solution)
        {
            var collectedMethods = new HashSet<IMethodSymbol>(SymbolEqualityComparer.Default);
            var methodsToProcess = new Queue<IMethodSymbol>();
            methodsToProcess.Enqueue(rootMethod);

            while (methodsToProcess.Count > 0)
            {
                var currentMethod = methodsToProcess.Dequeue();
                if (!collectedMethods.Add(currentMethod)) continue;

                try
                {
                    // Get implementations if this is an interface method
                    if (currentMethod.ContainingType?.TypeKind == TypeKind.Interface)
                    {
                        await ProcessInterfaceMethodAsync(currentMethod, solution, methodsToProcess, collectedMethods);
                        continue;
                    }

                    var methodSyntax = await GetMethodSyntaxAsync(currentMethod, solution);
                    if (methodSyntax == null) continue;

                    var document = solution.GetDocument(methodSyntax.SyntaxTree);
                    if (document == null) continue;

                    var semanticModel = await document.GetSemanticModelAsync();

                    // Process all invocation expressions
                    foreach (var invocation in methodSyntax.DescendantNodes().OfType<InvocationExpressionSyntax>())
                    {
                        var calledSymbol = await GetInvocationSymbolAsync(invocation, semanticModel);
                        if (calledSymbol == null) continue;

                        // Resolve interface methods to implementations
                        if (calledSymbol.ContainingType?.TypeKind == TypeKind.Interface)
                        {
                            var implementations = await FindImplementationsAsync(calledSymbol, solution);
                            foreach (var impl in implementations)
                            {
                                if (!collectedMethods.Contains(impl))
                                    methodsToProcess.Enqueue(impl);
                            }
                        }
                        else if (!collectedMethods.Contains(calledSymbol))
                        {
                            methodsToProcess.Enqueue(calledSymbol);
                        }
                    }
                }
                catch (Exception ex)
                {
                    await ex.LogAsync();
                    continue;
                }
            }

            return collectedMethods;
        }

        private static async Task ProcessInterfaceMethodAsync(
            IMethodSymbol interfaceMethod,
            Solution solution,
            Queue<IMethodSymbol> methodsToProcess,
            HashSet<IMethodSymbol> collectedMethods)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var implementations = await FindImplementationsAsync(interfaceMethod, solution);
            foreach (var impl in implementations)
            {
                if (!collectedMethods.Contains(impl))
                    methodsToProcess.Enqueue(impl);
            }
        }

        private static async Task<IEnumerable<IMethodSymbol>> FindImplementationsAsync(
            IMethodSymbol methodSymbol,
            Solution solution)
        {
            var implementations = new List<IMethodSymbol>();

            // Find all implementing types across the solution
            var implementingTypes = await SymbolFinder.FindImplementationsAsync(
                methodSymbol.ContainingType,
                solution,
                cancellationToken: default);

            foreach (var type in implementingTypes)
            {
                // Get the implementing method
                var implementation = type.FindImplementationForInterfaceMember(methodSymbol) as IMethodSymbol;
                if (implementation != null && !implementation.IsAbstract)
                {
                    implementations.Add(implementation);

                    // Include overrides in derived classes
                    var overrides = (await SymbolFinder.FindOverridesAsync(implementation, solution))
                        .OfType<IMethodSymbol>();
                    implementations.AddRange(overrides);
                }
            }

            return implementations.Distinct();
        }

        private static async Task<IMethodSymbol> GetInvocationSymbolAsync(
            InvocationExpressionSyntax invocation,
            SemanticModel semanticModel)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var symbolInfo = semanticModel.GetSymbolInfo(invocation);
            return symbolInfo.Symbol as IMethodSymbol
                ?? symbolInfo.CandidateSymbols.FirstOrDefault() as IMethodSymbol;
        }

        private static async Task<string> GenerateSourceWithDependenciesAsync1(
            IEnumerable<IMethodSymbol> methods, Solution solution)
        {
            var classes = methods
                .Where(m => m.ContainingType != null)
                .GroupBy(m => m.ContainingType)
                .OrderBy(g => g.Key?.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat))
                .ToList();

            var sb = new StringBuilder();
            var processedNamespaces = new HashSet<string>();
            var processedClasses = new HashSet<ITypeSymbol>();

            foreach (var classGroup in classes)
            {
                var typeSymbol = classGroup.Key;
                var methodsInClass = classGroup.Distinct().ToList();

                // Skip system namespace visualization for external types
                bool isExternal = typeSymbol.Locations.All(loc => loc.IsInMetadata);
                string ns = typeSymbol.ContainingNamespace?.ToDisplayString();

                if (!isExternal && !string.IsNullOrEmpty(ns) && processedNamespaces.Add(ns))
                {
                    sb.AppendLine($"namespace {ns}");
                    sb.AppendLine("{");
                }

                // Class/interface/struct declaration
                if (processedClasses.Add(typeSymbol))
                {
                    if (isExternal)
                    {
                        sb.AppendLine($"// External type from assembly: {typeSymbol.ContainingAssembly?.Name}");
                    }
                    sb.AppendLine($"    {GetTypeDeclaration(typeSymbol)}");
                    sb.AppendLine("    {");
                }

                foreach (var method in methodsInClass.OrderBy(m => m.Name))
                {
                    if (method.Locations.Any(loc => loc.IsInSource))
                    {
                        // Process source-available methods normally
                        var syntax = await GetMethodSyntaxAsync(method, solution);
                        if (syntax != null)
                        {
                            var formattedMethod = Formatter.Format(syntax, solution.Workspace)
                                .ToFullString()
                                .Replace("\n", "\n        ");

                            sb.AppendLine($"        {formattedMethod}");
                        }
                    }
                    else
                    {
                        // Generate signature for external methods
                        sb.AppendLine($"        {GetMethodSignature(method)}");
                        sb.AppendLine($"        {{ \n            throw new NotImplementedException(\"Original implementation in {method.ContainingAssembly?.Name}\"); \n        }}        ");
                    }
                    sb.AppendLine();
                }

                sb.AppendLine("    }");

                if (!isExternal && !string.IsNullOrEmpty(ns) &&
                    !classes.Any(c => c.Key.ContainingNamespace?.ToDisplayString() == ns &&
                                    !processedClasses.Contains(c.Key)))
                {
                    sb.AppendLine("}");
                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }




        private static async Task<string> GenerateSourceWithDependenciesAsync(
        IEnumerable<IMethodSymbol> methods, Solution solution)
        {
            var sb = new StringBuilder();
            var processedTypes = new HashSet<ITypeSymbol>();
            var processedNamespaces = new HashSet<string>();

            // Group by containing type
            var methodsByType = methods
                .Where(m => m.ContainingType != null)
                .GroupBy(m => m.ContainingType)
                .OrderBy(g => g.Key?.ToDisplayString());

            foreach (var typeGroup in methodsByType)
            {
                var typeSymbol = typeGroup.Key;
                var methodsInType = typeGroup.ToList();

                // Handle namespace
                if (typeSymbol.ContainingNamespace != null)
                {
                    var ns = typeSymbol.ContainingNamespace.ToDisplayString();
                    if (processedNamespaces.Add(ns))
                    {
                        sb.AppendLine($"namespace {ns}");
                        sb.AppendLine("{");
                    }
                }

                // Type declaration
                if (processedTypes.Add(typeSymbol))
                {
                    sb.AppendLine(GetTypeDeclaration(typeSymbol));
                    sb.AppendLine("{");
                }

                // Process methods
                foreach (var method in methodsInType.OrderBy(m => m.Name))
                {
                    sb.AppendLine(await GetMethodSourceAsync(method, solution));
                    sb.AppendLine();
                }

                sb.AppendLine("}");

                // Close namespace if needed
                if (typeSymbol.ContainingNamespace != null)
                {
                    var ns = typeSymbol.ContainingNamespace.ToDisplayString();
                    if (!methodsByType.Any(g => g.Key.ContainingNamespace?.ToDisplayString() == ns &&
                                              !processedTypes.Contains(g.Key)))
                    {
                        sb.AppendLine("}");
                        sb.AppendLine();
                    }
                }
            }

            return sb.ToString();
        }

        private static async Task<string> GetMethodSourceAsync(IMethodSymbol method, Solution solution)
        {
            // For VB.NET methods - try to decompile and convert
            if (method.Language == LanguageNames.VisualBasic)
            {
                try
                {
                    // Option 1: Try to get source via Roslyn first
                    if (method.DeclaringSyntaxReferences.Length > 0)
                    {
                        var syntaxRef = method.DeclaringSyntaxReferences[0];
                        var vbSyntax = await syntaxRef.GetSyntaxAsync();
                        var vbDocument = solution.GetDocument(syntaxRef.SyntaxTree);

                        // Use Roslyn's VB->C# conversion
                        var csharpCode = await ConvertVbToCSharpAsync(vbSyntax.ToFullString(), vbDocument);
                        return IndentMethodBody(csharpCode);
                    }

                    foreach (var project in solution.Projects)
                    {
                        var compilation = await project.GetCompilationAsync();
                        var decompiled = DecompileMethod(method, compilation);

                        if (!string.IsNullOrEmpty(decompiled))
                        {
                            var csharpCode = await ConvertVbToCSharpAsync(decompiled);
                            return IndentMethodBody(csharpCode);
                        }
                    }

                    // Option 2: Decompile from metadata if source not available
                }
                catch (Exception ex)
                {
                    await ex.LogAsync();
                    // Fall through to signature-only version
                }
            }

            // Fallback for C# methods or when conversion fails
            return await GetMethodSourceFallbackAsync(method, solution);
        }

        private static async Task<string> ConvertVbToCSharpAsync(string vbCode, Microsoft.CodeAnalysis.Document vbDocument = null)
        {
            var workspace = vbDocument?.Project.Solution.Workspace;

            if (workspace is null)
            {
                var componentModel = await VS.GetServiceAsync<SComponentModel, IComponentModel>();
                workspace = componentModel.GetService<VisualStudioWorkspace>();
            }

            string fromLanguage = LanguageNames.VisualBasic;
            string toLanguage = LanguageNames.CSharp;

            var codeWithOptions = new ICSharpCode.CodeConverter.CodeWithOptions(vbCode)
               .SetFromLanguage(fromLanguage)
               .SetToLanguage(toLanguage);

            var convertionResult = await ICSharpCode.CodeConverter.CodeConverter.Convert(codeWithOptions);

            if (convertionResult.Exceptions.Count > 0)
            {
                return vbCode;
            }

            //var conversionService = new CSharpConverter();
            //var converted = await conversionService.ConvertTextAsync(vbCode, workspace);
            return convertionResult.ConvertedCode;
        }
        private static string DecompileMethod(IMethodSymbol methodSymbol, Compilation compilation)
        {

            var assemblySymbol = methodSymbol.ContainingAssembly;

            string assemblyPath = null;

            foreach (var reference in compilation.References)
            {
                var symbol = compilation.GetAssemblyOrModuleSymbol(reference) as IAssemblySymbol;
                if (symbol != null && SymbolEqualityComparer.Default.Equals(symbol, assemblySymbol))
                {
                    if (reference is PortableExecutableReference peRef)
                    {
                        assemblyPath = peRef.FilePath;
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(assemblyPath)) return null;

            // 2. Setup Decompiler
            var peFile = new PEFile(assemblyPath);
            var resolver = new UniversalAssemblyResolver(assemblyPath, false, peFile.DetectTargetFrameworkId());
            var decompiler = new CSharpDecompiler(assemblyPath, resolver, new DecompilerSettings());

            // 3. Find the IMethod in the Decompiler TypeSystem
            var typeSystem = decompiler.TypeSystem;
            var mainModule = typeSystem.MainModule;

            var topLevelTypeName = new ICSharpCode.Decompiler.TypeSystem.TopLevelTypeName(
                methodSymbol.ContainingNamespace.ToDisplayString(),
                methodSymbol.ContainingType.Name
                );

            var type = mainModule.GetTypeDefinition(topLevelTypeName);


            if (type == null)
                throw new Exception("Type not found in PE file");

            // Match method by name + parameter count (you can extend this match as needed)
            var method = type.Methods.FirstOrDefault(m =>
                m.Name == methodSymbol.Name &&
                m.Parameters.Count == methodSymbol.Parameters.Length);

            if (method == null)
                throw new Exception("Method not found in PE file");

            // 4. Decompile that method
            return decompiler.DecompileAsString(method.MetadataToken);
        }


        private static string DecompileMethod1(IMethodSymbol method, Compilation compilation)
        {
            try
            {
                var assemblySymbol = method.ContainingAssembly;

                string assemblyPath = null;

                foreach (var reference in compilation.References)
                {
                    var symbol = compilation.GetAssemblyOrModuleSymbol(reference) as IAssemblySymbol;
                    if (symbol != null && SymbolEqualityComparer.Default.Equals(symbol, assemblySymbol))
                    {
                        if (reference is PortableExecutableReference peRef)
                        {
                            assemblyPath = peRef.FilePath;
                            break;
                        }
                    }
                }

                if (string.IsNullOrEmpty(assemblyPath)) return null;

                var decompiler = new CSharpDecompiler(assemblyPath,
                    new DecompilerSettings
                    {
                        ThrowOnAssemblyResolveErrors = false,
                        ShowXmlDocumentation = true
                    });

                var typeName = method.ContainingType.ToDisplayString(
                    SymbolDisplayFormat.FullyQualifiedFormat)
                    .Replace("global::", "");

                var typeCode = decompiler.DecompileTypeAsString(
                    new ICSharpCode.Decompiler.TypeSystem.FullTypeName(typeName));

                return ExtractMethodFromTypeCode(typeCode, method.Name);
            }
            catch
            {
                return null;
            }
        }

        private static string ExtractMethodFromTypeCode(string typeCode, string methodName)
        {
            // This is a simplified approach - consider using proper parsing for production
            var methodPattern = $@"(public|private|protected|internal)\s+.+\s+{methodName}\s*\([^)]*\)[^{{]*{{[^}}]*}}";
            var match = Regex.Match(typeCode, methodPattern, RegexOptions.Singleline);
            return match.Success ? match.Value : null;
        }

        //private static string DecompileMethod(IMethodSymbol method, Compilation compilation)
        //{
        //    try
        //    {
        //        // Use ILSpy to decompile
        //        var assemblyLocation = method.ContainingAssembly?.Locations.FirstOrDefault()?.FilePath;
        //        if (string.IsNullOrEmpty(assemblyLocation)) return null;

        //        var decompiler = new CSharpDecompiler(assemblyLocation,
        //            new DecompilerSettings
        //            {
        //                ThrowOnAssemblyResolveErrors = false,
        //                ShowXmlDocumentation = true
        //            });

        //        var typeName = method.ContainingType.ToDisplayString(
        //            SymbolDisplayFormat.FullyQualifiedFormat)
        //            .Replace("global::", "");

        //        var methodToken = method.GetSymbolToken();
        //        if (methodToken == 0) return null;

        //        return decompiler.DecompileAsString(typeName, methodToken);
        //    }
        //    catch
        //    {
        //        return null;
        //    }
        //}

        private static async Task<string> GetMethodSourceFallbackAsync(IMethodSymbol method, Solution solution)
        {
            // For C# methods or when conversion fails
            if (method.DeclaringSyntaxReferences.Length > 0)
            {
                var syntaxRef = method.DeclaringSyntaxReferences[0];
                var syntax = await syntaxRef.GetSyntaxAsync();
                var document = solution.GetDocument(syntaxRef.SyntaxTree);

                if (document != null)
                {
                    var formatted = Formatter.Format(syntax, document.Project.Solution.Workspace);
                    return IndentMethodBody(formatted.ToFullString());
                }
            }

            // Generate just the signature if no source available
            return GetMethodSignature(method);
        }

        private static string IndentMethodBody(string code)
        {
            return string.Join(Environment.NewLine,
                code.Split(new[] { Environment.NewLine }, StringSplitOptions.None)
                .Select(line => "    " + line));
        }





        private static string GetMethodSignature(IMethodSymbol method)
        {
            var accessibility = method.DeclaredAccessibility.ToString().ToLower();
            var returnType = method.ReturnType.ToDisplayString();
            var parameters = string.Join(", ", method.Parameters.Select(p =>
                $"{p.Type.ToDisplayString()} {p.Name}"));

            var modifiers = new List<string>();
            if (method.IsStatic) modifiers.Add("static");
            if (method.IsVirtual) modifiers.Add("virtual");
            if (method.IsAsync) modifiers.Add("async");

            return $"{accessibility} {string.Join(" ", modifiers)} {returnType} {method.Name}({parameters})";
        }

        private static string GetTypeDeclaration(ITypeSymbol typeSymbol)
        {
            var kind = typeSymbol.TypeKind switch
            {
                TypeKind.Class => "class",
                TypeKind.Struct => "struct",
                TypeKind.Interface => "interface",
                _ => "class"
            };

            var baseTypes = new List<string>();

            if (typeSymbol.BaseType != null &&
                typeSymbol.BaseType.SpecialType != SpecialType.System_Object)
            {
                baseTypes.Add(typeSymbol.BaseType.ToDisplayString());
            }

            baseTypes.AddRange(typeSymbol.Interfaces.Select(i => i.ToDisplayString()));

            var baseClause = baseTypes.Any()
                ? $" : {string.Join(", ", baseTypes)}"
                : string.Empty;

            return $"{kind} {typeSymbol.Name}{baseClause}";
        }

        private static async Task<SyntaxNode> GetMethodSyntaxAsync(IMethodSymbol method, Solution solution)
        {
            var reference = method.DeclaringSyntaxReferences.FirstOrDefault();
            if (reference == null) return null;

            return await reference.GetSyntaxAsync();
        }
    }

    public static class TextBufferExtensions
    {
        public static async Task<Microsoft.CodeAnalysis.Document> GetRoslynDocumentAsync(this ITextBuffer textBuffer)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (!textBuffer.Properties.TryGetProperty(typeof(ITextDocument), out ITextDocument textDoc))
                return null;

            var componentModel = await VS.GetServiceAsync<SComponentModel, IComponentModel>();
            var workspace = componentModel.GetService<VisualStudioWorkspace>();
            var documentId = workspace.CurrentSolution.GetDocumentIdsWithFilePath(textDoc.FilePath).FirstOrDefault();

            return documentId != null ? workspace.CurrentSolution.GetDocument(documentId) : null;
        }
    }
}