using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using EnvDTE;
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

        private static async Task<string> GenerateSourceWithDependenciesAsync(
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