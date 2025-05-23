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
/*
namespace MethodCopier
{
    [Command(PackageIds.MyCommand)]
    internal sealed class MyCommand : BaseCommand<MyCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            try
            {
                var doc = await VS.Documents.GetActiveDocumentViewAsync();
                if (doc?.TextView == null) return;

                var selection = doc.TextView.Selection.SelectedSpans.FirstOrDefault();
                if (selection.IsEmpty) return;

                var textSpan = new TextSpan(selection.Start, selection.Length);
                
                var roslynDoc = await GetActiveRoslynDocumentAsync();
                if (roslynDoc == null) return;

                var model = await roslynDoc.GetSemanticModelAsync();
                var root = await roslynDoc.GetSyntaxRootAsync();
                var node = root.FindNode(textSpan);

                var methodSymbol = await GetMethodSymbolAsync(model, node);
                if (methodSymbol == null) return;

                var dependencies = await FindMethodDependenciesAsync(methodSymbol, roslynDoc.Project.Solution);
                var allMethods = dependencies.Append(methodSymbol).Distinct();

                var sourceText = await GenerateSourceWithDependenciesAsync(allMethods, roslynDoc.Project.Solution);
                
                Clipboard.SetText(sourceText);

                await VS.StatusBar.ShowMessageAsync("Method and dependencies copied to clipboard!");
            }
            catch (Exception ex)
            {
                await ex.LogAsync();
                await VS.MessageBox.ShowErrorAsync("Error copying method", ex.Message);
            }
        }
        public static async Task<Microsoft.CodeAnalysis.Document?> GetActiveRoslynDocumentAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();


            DTE2 dte = (DTE2)await VS.GetServiceAsync<DTE, DTE>();

            //DTE2 dte = this.ServiceProvider.GetService(typeof(DTE)) as DTE2;
            FileCodeModel fcm = dte.ActiveDocument.ProjectItem.FileCodeModel as FileCodeModel;
            foreach (CodeElement element in fcm.CodeElements)
            {
                if (element is CodeNamespace)
                {
                    CodeNamespace nsp = element as CodeNamespace;

                    foreach (CodeElement subElement in nsp.Children)
                    {
                        if (subElement is CodeClass)
                        {
                            CodeClass c2 = subElement as CodeClass;
                            foreach (CodeElement item in c2.Children)
                            {
                                if (item is CodeFunction)
                                {
                                    CodeFunction cf = item as CodeFunction;
                                }
                            }
                        }
                    }
                }
            }







            // Get the active document via DTE (VS automation)
            var dte1 = await VS.GetServiceAsync<DTE, DTE>();
            var activeDoc = dte1.ActiveDocument;

            if (activeDoc == null || string.IsNullOrEmpty(activeDoc.FullName))
                return null;

            // Read the file content manually
            string filePath = activeDoc.FullName;
            string fileText = File.ReadAllText(filePath);

            // Create an Ad-Hoc Roslyn Workspace
            var workspace = new AdhocWorkspace();
            var project = workspace.AddProject("TempProject", LanguageNames.CSharp);
            var document = workspace.AddDocument(project.Id, Path.GetFileName(filePath), SourceText.From(fileText));

            return document;
        }

        private static async Task<string> GenerateSourceWithDependenciesAsync(IEnumerable<IMethodSymbol> methods, Microsoft.CodeAnalysis.Solution solution)
        {
            var sb = new StringBuilder();

            foreach (var method in methods)
            {
                var location = method.Locations.FirstOrDefault(l => l.IsInSource);
                if (location == null) continue;

                var doc = solution.GetDocument(location.SourceTree);
                if (doc == null) continue;

                var root = await doc.GetSyntaxRootAsync();
                var node = root.FindNode(location.SourceSpan);
                sb.AppendLine(node.ToFullString());
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static async Task<IEnumerable<IMethodSymbol>> FindMethodDependenciesAsync(IMethodSymbol method, Microsoft.CodeAnalysis.Solution solution)
        {
            var dependencies = new HashSet<IMethodSymbol>();
            var toProcess = new Queue<IMethodSymbol>();
            toProcess.Enqueue(method);

            while (toProcess.Count > 0)
            {
                var current = toProcess.Dequeue();
                var references = await SymbolFinder.FindReferencesAsync(current, solution);

                foreach (var reference in references)
                {
                    foreach (var location in reference.Locations)
                    {
                        var doc = solution.GetDocument(location.Document.Id);
                        if (doc == null) continue;

                        var model = await doc.GetSemanticModelAsync();
                        var root = await doc.GetSyntaxRootAsync();
                        var node = root.FindNode(location.Location.SourceSpan);

                        var containingMethod = node.AncestorsAndSelf()
                            .Select(n => model.GetDeclaredSymbol(n))
                            .FirstOrDefault(s => s is IMethodSymbol) as IMethodSymbol;

                        if (containingMethod != null && dependencies.Add(containingMethod))
                        {
                            toProcess.Enqueue(containingMethod);
                        }
                    }
                }
            }

            return dependencies;
        }

        private static async Task<IMethodSymbol> GetMethodSymbolAsync(SemanticModel model, SyntaxNode node)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            // Check if we're on a method declaration
            var declaredSymbol = model.GetDeclaredSymbol(node) as IMethodSymbol;
            if (declaredSymbol != null) return declaredSymbol;

            // Check if we're on a method reference
            var symbolInfo = model.GetSymbolInfo(node);
            return symbolInfo.Symbol as IMethodSymbol ?? symbolInfo.CandidateSymbols.FirstOrDefault() as IMethodSymbol;
        }
    }
}

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
                var dte = await VS.GetServiceAsync<DTE, DTE2>();
                var activeDoc = dte.ActiveDocument;

                if (activeDoc == null)
                {
                    await VS.MessageBox.ShowErrorAsync("No active document");
                    return;
                }

                var selection = (TextSelection)dte.ActiveDocument.Selection;
                var fcm = activeDoc.ProjectItem.FileCodeModel;

                var methods = fcm.CodeElements
                    .GetAllMethods()
                    .OfType<CodeFunction>()
                    .ToList();
                //    new List<CodeFunction>();
                //FindMethodsInSelection(fcm.CodeElements, selection, methods);

                if (methods.Count == 0)
                {
                    await VS.MessageBox.ShowErrorAsync("No method selected");
                    return;
                }

                var sourceCode = BuildMethodWithDependencies(methods.First());
                Clipboard.SetText(sourceCode);

                await VS.StatusBar.ShowMessageAsync("Method copied to clipboard!");
            }
            catch (Exception ex)
            {
                await VS.MessageBox.ShowErrorAsync("Error", ex.ToString());
            }
        }

        private void FindMethodsInSelection(CodeElements elements, TextSelection selection, List<CodeFunction> results)
        {
            foreach (CodeElement element in elements)
            {
                if (element is CodeNamespace ns)
                {
                    FindMethodsInSelection(ns.Children, selection, results);
                }
                else if (element is CodeClass cls)
                {
                    FindMethodsInSelection(cls.Children, selection, results);
                }
                else if (element is CodeFunction func)
                {
                    if (selection.TopPoint.Line >= func.StartPoint.Line &&
                        selection.BottomPoint.Line <= func.EndPoint.Line)
                    {
                        results.Add(func);
                    }
                }
            }
        }

        private string BuildMethodWithDependencies(CodeFunction mainMethod)
        {
            var sb = new StringBuilder();
            var processedMethods = new HashSet<string>();

            void AddMethod(CodeFunction method)
            {
                if (processedMethods.Contains(method.FullName)) return;

                sb.AppendLine($"{method.Access} {method.Type.AsString} {method.Name}({GetParameters(method)})");
                sb.AppendLine(method.StartPoint.CreateEditPoint().GetText(method.EndPoint));
                sb.AppendLine();

                processedMethods.Add(method.FullName);

                // Find called methods (simplified - would need proper parsing)
                var body = method.StartPoint.CreateEditPoint().GetText(method.EndPoint);
                FindCalledMethods(body, method.ProjectItem.FileCodeModel, processedMethods, AddMethod);
            }

            AddMethod(mainMethod);
            return sb.ToString();
        }

        private string GetParameters(CodeFunction method)
        {
            return string.Join(", ", method.Parameters.OfType<CodeParameter>()
                .Select(p => $"{p.Type.AsString} {p.Name}"));
        }

        private void FindCalledMethods(string body, FileCodeModel fcm, HashSet<string> processed, Action<CodeFunction> callback)
        {
            var allMethods = fcm.CodeElements
                .RecursiveSelect<CodeFunction>(e =>
                {
                    switch (e)
                    {
                        case CodeNamespace ns: return ns.Children;
                        case CodeClass cls: return cls.Children;
                        case CodeStruct st: return st.Children;
                        case CodeInterface iface: return iface.Children;
                        default: return null;
                    }
                });

            foreach (var method in allMethods)
            {
                if (body.Contains(method.Name + "(") && !processed.Contains(method.FullName))
                {
                    callback(method);
                }
            }
        }
    }

    // Helper extension
    public static class EnumerableExtensions
    {
        public static IEnumerable<T> RecursiveSelect<T>(this IEnumerable collection,
            Func<object, IEnumerable> childSelector)
            where T : class
        {
            if (collection == null) yield break;

            foreach (var item in collection)
            {
                if (item is T matchedItem)
                {
                    yield return matchedItem;
                }

                var children = childSelector(item);
                if (children != null)
                {
                    foreach (var child in children.RecursiveSelect<T>(childSelector))
                    {
                        yield return child;
                    }
                }
            }
        }

        public static IEnumerable<CodeFunction> GetAllMethods(this CodeElements elements)
        {
            if (elements == null) yield break;

            foreach (CodeElement element in elements)
            {
                if (element is CodeFunction function)
                {
                    yield return function;
                }

                CodeElements children = element switch
                {
                    CodeNamespace ns => ns.Children,
                    CodeClass cls => cls.Children,
                    CodeStruct st => st.Children,
                    CodeInterface iface => iface.Children,
                    _ => null
                };

                if (children != null)
                {
                    foreach (var childMethod in children.GetAllMethods())
                    {
                        yield return childMethod;
                    }
                }
            }
        }
    }
}
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
                var dte = await VS.GetServiceAsync<DTE, DTE2>();
                var activeDoc = dte.ActiveDocument;

                if (activeDoc == null)
                {
                    await VS.MessageBox.ShowErrorAsync("No active document");
                    return;
                }

                var selection = (TextSelection)dte.ActiveDocument.Selection;
                var fcm = activeDoc.ProjectItem.FileCodeModel;

                // Find the initially selected method
                var selectedMethod = FindMethodAtSelection(fcm.CodeElements, selection);
                if (selectedMethod == null)
                {
                    await VS.MessageBox.ShowErrorAsync("No method selected");
                    return;
                }

                // Get all methods in the file
                var allMethods = GetAllMethods(fcm.CodeElements).ToList();

                // Find dependencies recursively
                var methodsToCopy = new List<CodeFunction>();
                FindMethodDependencies(selectedMethod, allMethods, methodsToCopy);

                // Generate source code
                var sourceCode = GenerateSourceCode(methodsToCopy);
                Clipboard.SetText(sourceCode);

                await VS.StatusBar.ShowMessageAsync($"Copied {methodsToCopy.Count} methods to clipboard!");
            }
            catch (Exception ex)
            {
                await VS.MessageBox.ShowErrorAsync("Error", ex.ToString());
            }
        }

        private CodeFunction FindMethodAtSelection(CodeElements elements, TextSelection selection)
        {
            foreach (CodeElement element in elements)
            {
                if (element is CodeNamespace ns)
                {
                    var result = FindMethodAtSelection(ns.Children, selection);
                    if (result != null) return result;
                }
                else if (element is CodeClass cls)
                {
                    var result = FindMethodAtSelection(cls.Children, selection);
                    if (result != null) return result;
                }
                else if (element is CodeFunction func)
                {
                    if (selection.TopPoint.Line >= func.StartPoint.Line &&
                        selection.BottomPoint.Line <= func.EndPoint.Line)
                    {
                        return func;
                    }
                }
            }
            return null;
        }

        private IEnumerable<CodeFunction> GetAllMethods(CodeElements elements)
        {
            if (elements == null) yield break;

            foreach (CodeElement element in elements)
            {
                if (element is CodeFunction function)
                {
                    yield return function;
                }

                CodeElements children = element switch
                {
                    CodeNamespace ns => ns.Children,
                    CodeClass cls => cls.Children,
                    CodeStruct st => st.Children,
                    CodeInterface iface => iface.Children,
                    _ => null
                };

                if (children != null)
                {
                    foreach (var childMethod in GetAllMethods(children))
                    {
                        yield return childMethod;
                    }
                }
            }
        }

        private void FindMethodDependencies(CodeFunction method, List<CodeFunction> allMethods, List<CodeFunction> result)
        {
            if (result.Contains(method)) return;

            result.Add(method);
            var methodBody = method.StartPoint.CreateEditPoint().GetText(method.EndPoint);

            foreach (var potentialDependency in allMethods)
            {
                // Simple check - would need enhancement for production use
                if (methodBody.Contains(potentialDependency.Name + "(") &&
                    !result.Contains(potentialDependency))
                {
                    FindMethodDependencies(potentialDependency, allMethods, result);
                }
            }
        }

        private string GenerateSourceCode(List<CodeFunction> methods)
        {
            var sb = new StringBuilder();
            foreach (var method in methods)
            {
                sb.AppendLine($"{method.Access} {method.Type.AsString} {method.Name}({GetParameters(method)})");
                sb.AppendLine(method.StartPoint.CreateEditPoint().GetText(method.EndPoint));
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private string GetParameters(CodeFunction method)
        {
            return string.Join(", ", method.Parameters.OfType<CodeParameter>()
                .Select(p => $"{p.Type.AsString} {p.Name}"));
        }
    }
}

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
            IMethodSymbol rootMethod, Microsoft.CodeAnalysis.Solution solution)
        {
            var collectedMethods = new HashSet<IMethodSymbol>(SymbolEqualityComparer.Default);
            var methodsToProcess = new Queue<IMethodSymbol>();
            methodsToProcess.Enqueue(rootMethod);

            while (methodsToProcess.Count > 0)
            {
                var currentMethod = methodsToProcess.Dequeue();
                if (!collectedMethods.Add(currentMethod)) continue;

                // Get method body syntax
                var methodSyntax = await GetMethodSyntaxAsync(currentMethod, solution);
                if (methodSyntax == null) continue;

                // Find all method calls in the body
                var methodCalls = methodSyntax.DescendantNodes()
                    .OfType<InvocationExpressionSyntax>()
                    .Select(i => i.Expression);

                foreach (var call in methodCalls)
                {
                    var calledSymbol = solution.GetDocument(call.SyntaxTree)
                        ?.GetSemanticModelAsync().Result
                        ?.GetSymbolInfo(call).Symbol as IMethodSymbol;

                    if (calledSymbol != null && !collectedMethods.Contains(calledSymbol))
                        methodsToProcess.Enqueue(calledSymbol);
                }

                // Also process property getters/setters
                var propertyAccess = methodSyntax.DescendantNodes()
                    .OfType<MemberAccessExpressionSyntax>()
                    .Select(m => m.Name);

                foreach (var access in propertyAccess)
                {
                    var propertySymbol = solution.GetDocument(access.SyntaxTree)
                        ?.GetSemanticModelAsync().Result
                        ?.GetSymbolInfo(access).Symbol as IPropertySymbol;

                    if (propertySymbol != null)
                    {
                        if (propertySymbol.GetMethod != null && !collectedMethods.Contains(propertySymbol.GetMethod))
                            methodsToProcess.Enqueue(propertySymbol.GetMethod);
                        if (propertySymbol.SetMethod != null && !collectedMethods.Contains(propertySymbol.SetMethod))
                            methodsToProcess.Enqueue(propertySymbol.SetMethod);
                    }
                }
            }

            return collectedMethods;
        }

        private static async Task<SyntaxNode> GetMethodSyntaxAsync(IMethodSymbol method, Microsoft.CodeAnalysis.Solution solution)
        {
            var reference = method.DeclaringSyntaxReferences.FirstOrDefault();
            if (reference == null) return null;

            var syntaxTree = await reference.SyntaxTree.GetRootAsync();
            return reference.GetSyntax();
        }

        private static async Task<string> GenerateSourceWithDependenciesAsync(
            IEnumerable<IMethodSymbol> methods, Microsoft.CodeAnalysis.Solution solution)
        {
            var sb = new StringBuilder();

            foreach (var method in methods)
            {
                var syntaxReference = method.DeclaringSyntaxReferences.FirstOrDefault();
                if (syntaxReference == null) continue;

                var syntaxNode = await syntaxReference.GetSyntaxAsync();
                var doc = solution.GetDocument(syntaxReference.SyntaxTree);
                var formattedNode = await Formatter.FormatAsync(syntaxNode, doc);

                sb.AppendLine(formattedNode.ToFullString());
                sb.AppendLine();
            }

            return sb.ToString();
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

            return workspace.CurrentSolution.GetDocumentIdsWithFilePath(textDoc.FilePath)
                          .FirstOrDefault()
                          ?.GetDocument();
        }
    }
}

*/

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

        private static async Task<HashSet<IMethodSymbol>> FindAllMethodDependenciesAsync1(
            IMethodSymbol rootMethod, Microsoft.CodeAnalysis.Solution solution)
        {
            var collectedMethods = new HashSet<IMethodSymbol>(SymbolEqualityComparer.Default);
            var methodsToProcess = new Queue<IMethodSymbol>();
            methodsToProcess.Enqueue(rootMethod);

            while (methodsToProcess.Count > 0)
            {
                var currentMethod = methodsToProcess.Dequeue();
                if (!collectedMethods.Add(currentMethod)) continue;

                // Get method body syntax
                var methodSyntax = await GetMethodSyntaxAsync(currentMethod, solution);
                if (methodSyntax == null) continue;

                // Find all method calls in the body
                var methodCalls = methodSyntax.DescendantNodes()
                    .OfType<InvocationExpressionSyntax>()
                    .Select(i => i.Expression);

                foreach (var call in methodCalls)
                {
                    var syntaxTree = call.SyntaxTree;
                    var document = solution.GetDocument(syntaxTree);
                    if (document == null) continue;

                    var semanticModel = await document.GetSemanticModelAsync();
                    var calledSymbol = semanticModel.GetSymbolInfo(call).Symbol as IMethodSymbol;

                    if (calledSymbol != null && !collectedMethods.Contains(calledSymbol))
                        methodsToProcess.Enqueue(calledSymbol);
                }

                // Process property accessors
                var propertyAccess = methodSyntax.DescendantNodes()
                    .OfType<MemberAccessExpressionSyntax>()
                    .Select(m => m.Name);

                foreach (var access in propertyAccess)
                {
                    var syntaxTree = access.SyntaxTree;
                    var document = solution.GetDocument(syntaxTree);
                    if (document == null) continue;

                    var semanticModel = await document.GetSemanticModelAsync();
                    var propertySymbol = semanticModel.GetSymbolInfo(access).Symbol as IPropertySymbol;

                    if (propertySymbol != null)
                    {
                        if (propertySymbol.GetMethod != null && !collectedMethods.Contains(propertySymbol.GetMethod))
                            methodsToProcess.Enqueue(propertySymbol.GetMethod);
                        if (propertySymbol.SetMethod != null && !collectedMethods.Contains(propertySymbol.SetMethod))
                            methodsToProcess.Enqueue(propertySymbol.SetMethod);
                    }
                }
            }

            return collectedMethods;
        }




        private static async Task<HashSet<IMethodSymbol>> FindAllMethodDependenciesAsync2(
                IMethodSymbol rootMethod, Microsoft.CodeAnalysis.Solution solution)
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
                    var methodSyntax = await GetMethodSyntaxAsync(currentMethod, solution);
                    if (methodSyntax == null) continue;

                    // Find all invocations in the method body
                    var invocations = methodSyntax.DescendantNodes()
                        .OfType<InvocationExpressionSyntax>();

                    foreach (var invocation in invocations)
                    {
                        var calledSymbol = await GetInvocationSymbolAsync(invocation, solution);
                        if (calledSymbol != null && !collectedMethods.Contains(calledSymbol))
                        {
                            methodsToProcess.Enqueue(calledSymbol);
                        }
                    }

                    // Handle property accessors
                    var memberAccesses = methodSyntax.DescendantNodes()
                        .OfType<MemberAccessExpressionSyntax>();

                    foreach (var access in memberAccesses)
                    {
                        var propertySymbol = await GetMemberAccessSymbolAsync(access, solution);
                        if (propertySymbol != null)
                        {
                            if (propertySymbol.GetMethod != null && !collectedMethods.Contains(propertySymbol.GetMethod))
                                methodsToProcess.Enqueue(propertySymbol.GetMethod);
                            if (propertySymbol.SetMethod != null && !collectedMethods.Contains(propertySymbol.SetMethod))
                                methodsToProcess.Enqueue(propertySymbol.SetMethod);
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


        private static async Task<IMethodSymbol> GetInvocationSymbolAsync(
            InvocationExpressionSyntax invocation, Microsoft.CodeAnalysis.Solution solution)
        {
            var expression = invocation.Expression;
            var document = solution.GetDocument(invocation.SyntaxTree);
            if (document == null) return null;

            try
            {
                var semanticModel = await document.GetSemanticModelAsync();
                var symbolInfo = semanticModel.GetSymbolInfo(expression);

                return symbolInfo.Symbol as IMethodSymbol
                    ?? symbolInfo.CandidateSymbols.FirstOrDefault() as IMethodSymbol;
            }
            catch
            {
                return null;
            }
        }




        private static async Task<HashSet<IMethodSymbol>> FindAllMethodDependenciesAsync(
            IMethodSymbol rootMethod, Microsoft.CodeAnalysis.Solution solution)
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
                        await ProcessInterfaceMethod(currentMethod, solution, methodsToProcess, collectedMethods);
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

        private static async Task ProcessInterfaceMethod(
            IMethodSymbol interfaceMethod,
            Microsoft.CodeAnalysis.Solution solution,
            Queue<IMethodSymbol> methodsToProcess,
            HashSet<IMethodSymbol> collectedMethods)
        {
            var implementations = await FindImplementationsAsync(interfaceMethod, solution);
            foreach (var impl in implementations)
            {
                if (!collectedMethods.Contains(impl))
                    methodsToProcess.Enqueue(impl);
            }
        }

        private static async Task<IEnumerable<IMethodSymbol>> FindImplementationsAsync(
            IMethodSymbol methodSymbol,
            Microsoft.CodeAnalysis.Solution solution)
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
            var symbolInfo = semanticModel.GetSymbolInfo(invocation);
            return symbolInfo.Symbol as IMethodSymbol
                ?? symbolInfo.CandidateSymbols.FirstOrDefault() as IMethodSymbol;
        }












        private static async Task ProcessSymbolInfo(
            SymbolInfo symbolInfo,
            Queue<IMethodSymbol> methodsToProcess,
            HashSet<IMethodSymbol> collectedMethods,
            Microsoft.CodeAnalysis.Solution solution)
        {
            var symbol = symbolInfo.Symbol ?? symbolInfo.CandidateSymbols.FirstOrDefault();

            if (symbol is IMethodSymbol methodSymbol)
            {
                if (!collectedMethods.Contains(methodSymbol))
                    methodsToProcess.Enqueue(methodSymbol);
            }
            else if (symbol is IPropertySymbol propertySymbol)
            {
                if (propertySymbol.GetMethod != null && !collectedMethods.Contains(propertySymbol.GetMethod))
                    methodsToProcess.Enqueue(propertySymbol.GetMethod);
                if (propertySymbol.SetMethod != null && !collectedMethods.Contains(propertySymbol.SetMethod))
                    methodsToProcess.Enqueue(propertySymbol.SetMethod);
            }
            else if (symbol is IFieldSymbol fieldSymbol && fieldSymbol.AssociatedSymbol != null)
            {
                // Handle field-like events
                if (fieldSymbol.AssociatedSymbol is IMethodSymbol associatedMethod)
                {
                    if (!collectedMethods.Contains(associatedMethod))
                        methodsToProcess.Enqueue(associatedMethod);
                }
            }
        }

        private static async Task<IMethodSymbol> GetInvocationSymbolAsync1(
            InvocationExpressionSyntax invocation,
            SemanticModel semanticModel)
        {
            try
            {
                var symbolInfo = semanticModel.GetSymbolInfo(invocation);

                // Handle extension methods
                if (symbolInfo.Symbol == null && symbolInfo.CandidateSymbols.Length > 0)
                {
                    return symbolInfo.CandidateSymbols
                        .OfType<IMethodSymbol>()
                        .FirstOrDefault(m => m.ReducedFrom != null);
                }

                return symbolInfo.Symbol as IMethodSymbol
                    ?? symbolInfo.CandidateSymbols.FirstOrDefault() as IMethodSymbol;
            }
            catch
            {
                return null;
            }
        }

        private static async Task<SyntaxNode> GetMethodSyntaxAsync2(IMethodSymbol method, Microsoft.CodeAnalysis.Solution solution)
        {
            if (method.DeclaringSyntaxReferences.IsEmpty)
                return null;

            var reference = method.DeclaringSyntaxReferences.First();
            var document = solution.GetDocument(reference.SyntaxTree);

            if (document == null)
                return null;

            return await reference.GetSyntaxAsync();
        }







        private static async Task<IPropertySymbol> GetMemberAccessSymbolAsync(
            MemberAccessExpressionSyntax memberAccess, Microsoft.CodeAnalysis.Solution solution)
        {
            var document = solution.GetDocument(memberAccess.SyntaxTree);
            if (document == null) return null;

            try
            {
                var semanticModel = await document.GetSemanticModelAsync();
                var symbolInfo = semanticModel.GetSymbolInfo(memberAccess.Name);

                return symbolInfo.Symbol as IPropertySymbol
                    ?? symbolInfo.CandidateSymbols.FirstOrDefault() as IPropertySymbol;
            }
            catch
            {
                return null;
            }
        }





        private static async Task<SyntaxNode> GetMethodSyntaxAsync1(IMethodSymbol method, Microsoft.CodeAnalysis.Solution solution)
        {
            var reference = method.DeclaringSyntaxReferences.FirstOrDefault();
            if (reference == null) return null;

            return await reference.GetSyntaxAsync();
        }

        private static async Task<string> GenerateSourceWithDependenciesAsync(
    IEnumerable<IMethodSymbol> methods, Microsoft.CodeAnalysis.Solution solution)
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
                        sb.AppendLine($"        {{ \n    throw new NotImplementedException(\"Original implementation in {method.ContainingAssembly?.Name}\"); \n }}        ");
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

private static async Task<string> GenerateSourceWithDependenciesAsync4(
    IEnumerable<IMethodSymbol> methods, Microsoft.CodeAnalysis.Solution solution)
        {
            var classes = methods
                .Where(m => m.ContainingType != null)
                .GroupBy(m => m.ContainingType)
                .OrderBy(g => g.Key?.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat))
                .ToList();

            var sb = new StringBuilder();
            var processedNamespaces = new HashSet<string>();
            var processedClasses = new HashSet<ITypeSymbol>();

            // Initialize progress
            int totalClasses = classes.Count;
            int processedClassesCount = 0;

            foreach (var classGroup in classes)
            {
                processedClassesCount++;
                var typeSymbol = classGroup.Key;
                var methodsInClass = classGroup.Distinct().ToList();

                // Update progress - file level
                await VS.StatusBar.ShowProgressAsync(
                    $"Processing {typeSymbol.Name}",
                    processedClassesCount,
                    totalClasses);

                // Process namespace
                if (typeSymbol.ContainingNamespace != null)
                {
                    var ns = typeSymbol.ContainingNamespace.ToDisplayString();
                    if (processedNamespaces.Add(ns))
                    {
                        sb.AppendLine($"namespace {ns}");
                        sb.AppendLine("{");
                    }
                }

                // Process class
                if (processedClasses.Add(typeSymbol))
                {
                    sb.AppendLine($"    {GetTypeDeclaration(typeSymbol)}");
                    sb.AppendLine("    {");
                }

                // Process methods with method-level progress
                int totalMethods = methodsInClass.Count;
                int processedMethods = 0;

                foreach (var method in methodsInClass.OrderBy(m => m.Name))
                {
                    processedMethods++;

                    // Only update progress every N items to reduce UI overhead
                    if (processedClassesCount % 5 == 0 || processedClassesCount == totalClasses)
                    {
                        await VS.StatusBar.ShowProgressAsync(
                            $"Processing {typeSymbol.Name}.{method.Name}",
                            processedMethods,
                            totalMethods);
                    }

                    var syntax = await GetMethodSyntaxAsync(method, solution);
                    if (syntax != null)
                    {
                        var formattedMethod = Formatter.Format(syntax, solution.Workspace)
                            .ToFullString()
                            .Replace("\n", "\n        ");

                        sb.AppendLine($"        {formattedMethod}");
                        sb.AppendLine();
                    }
                }

                // Close class
                sb.AppendLine("    }");

                // Close namespace if done
                if (typeSymbol.ContainingNamespace != null)
                {
                    var ns = typeSymbol.ContainingNamespace.ToDisplayString();
                    if (!classes.Any(c => c.Key.ContainingNamespace?.ToDisplayString() == ns &&
                                        !processedClasses.Contains(c.Key)))
                    {
                        sb.AppendLine("}");
                        sb.AppendLine();
                    }
                }
            }

            await VS.StatusBar.EndAnimationAsync(StatusAnimation.Save);
            return sb.ToString();
        }

        private static async Task<string> GenerateSourceWithDependenciesAsync3(
            IEnumerable<IMethodSymbol> methods, Microsoft.CodeAnalysis.Solution solution)
        {
            // Group methods by their containing class
            var classes = methods
                .Where(m => m.ContainingType != null)
                .GroupBy(m => m.ContainingType)
                .OrderBy(g => g.Key?.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat));

            var sb = new StringBuilder();

            // Track processed namespaces and classes
            var processedNamespaces = new HashSet<string>();
            var processedClasses = new HashSet<ITypeSymbol>();

            foreach (var classGroup in classes)
            {
                var typeSymbol = classGroup.Key;
                var methodsInClass = classGroup.ToList();

                // Add namespace declaration if not already added
                if (typeSymbol.ContainingNamespace != null)
                {
                    var ns = typeSymbol.ContainingNamespace.ToDisplayString();
                    if (processedNamespaces.Add(ns))
                    {
                        sb.AppendLine($"namespace {ns}");
                        sb.AppendLine("{");
                    }
                }

                // Add class declaration
                if (processedClasses.Add(typeSymbol))
                {
                    sb.AppendLine($"    {GetTypeDeclaration(typeSymbol)}");
                    sb.AppendLine("    {");
                }

                // Add all methods for this class
                foreach (var method in methodsInClass.OrderBy(m => m.Name))
                {
                    var syntax = await GetMethodSyntaxAsync(method, solution);
                    if (syntax != null)
                    {
                        var formattedMethod = Formatter.Format(syntax, solution.Workspace)
                            .ToFullString()
                            .Replace("\n", "\n        "); // Indent method body

                        sb.AppendLine($"        {formattedMethod}");
                        sb.AppendLine();
                    }
                }

                // Close class brace
                sb.AppendLine("    }");

                // Close namespace brace if no more classes in this namespace
                if (typeSymbol.ContainingNamespace != null)
                {
                    var ns = typeSymbol.ContainingNamespace.ToDisplayString();
                    if (!classes.Any(c => c.Key.ContainingNamespace?.ToDisplayString() == ns &&
                                        !processedClasses.Contains(c.Key)))
                    {
                        sb.AppendLine("}");
                        sb.AppendLine();
                    }
                }
            }

            return sb.ToString();
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

        private static async Task<SyntaxNode> GetMethodSyntaxAsync(IMethodSymbol method, Microsoft.CodeAnalysis.Solution solution)
        {
            var reference = method.DeclaringSyntaxReferences.FirstOrDefault();
            if (reference == null) return null;

            return await reference.GetSyntaxAsync();
        }







        private static async Task<string> GenerateSourceWithDependenciesAsync2(
            IEnumerable<IMethodSymbol> methods, Microsoft.CodeAnalysis.Solution solution)
        {
            var sb = new StringBuilder();
            sb.AppendLine("// =============================================");
            sb.AppendLine("// COPIED METHODS WITH DEPENDENCIES");
            sb.AppendLine($"// Generated on {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("// =============================================");
            sb.AppendLine();

            foreach (var method in methods.OrderBy(m => m.ContainingType?.Name).ThenBy(m => m.Name))
            {
                var syntaxReference = method.DeclaringSyntaxReferences.FirstOrDefault();
                if (syntaxReference == null) continue;

                var syntaxNode = await syntaxReference.GetSyntaxAsync();
                var document = solution.GetDocument(syntaxReference.SyntaxTree);

                if (document == null) continue;

                // Get source file information
                var filePath = document.FilePath;
                var fileName = Path.GetFileName(filePath);
                var lineSpan = syntaxReference.SyntaxTree.GetLineSpan(syntaxReference.Span);
                var lineNumber = lineSpan.StartLinePosition.Line + 1;

                // Add header comment
                sb.AppendLine("// =============================================");
                sb.AppendLine($"// Method: {method.ContainingType?.Name}.{method.Name}");
                sb.AppendLine($"// File: {fileName}");
                sb.AppendLine($"// Location: Line {lineNumber}");
                sb.AppendLine($"// Project: {document.Project.Name}");
                sb.AppendLine("// =============================================");

                // Add the actual method code
                var formattedNode = Formatter.Format(syntaxNode, solution.Workspace);
                sb.AppendLine(formattedNode.ToFullString());
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static async Task<string> GenerateSourceWithDependenciesAsync1(
            IEnumerable<IMethodSymbol> methods, Microsoft.CodeAnalysis.Solution solution)
        {
            var sb = new StringBuilder();
            var workspace = solution.Workspace;
            var options = workspace.Options;

            foreach (var method in methods)
            {
                var syntaxReference = method.DeclaringSyntaxReferences.FirstOrDefault();
                if (syntaxReference == null) continue;

                var syntaxNode = await syntaxReference.GetSyntaxAsync();
                var formattedNode = Formatter.Format(syntaxNode, workspace, options);

                sb.AppendLine(formattedNode.ToFullString());
                sb.AppendLine();
            }

            return sb.ToString();
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