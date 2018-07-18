Option Explicit On
Option Infer Off
Option Strict On
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.Simplification
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports AttributeListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.AttributeListSyntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports StatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.StatementSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)

            ' This file contains all the stuff accessed by multiple Visitor functions in Class NodeVisitor and Visitors that
            ' had no better home.

            ReadOnly allImports As List(Of ImportsStatementSyntax) = New List(Of ImportsStatementSyntax)()
            ReadOnly inlineAssignHelperMarkers As List(Of CSS.BaseTypeDeclarationSyntax) = New List(Of CSS.BaseTypeDeclarationSyntax)()
            Private _IsModule As Boolean = False
            Private IsModuleStack As New Stack(Of Boolean)
            Private mOptions As VisualBasicCompilationOptions
            Private mSemanticModel As SemanticModel
            Private mTargetDocument As Document
            Private NothingExpression As Syntax.LiteralExpressionSyntax = SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(SyntaxKind.NothingKeyword))
            Private placeholder As Integer = 1
            Public Sub New(ByVal lSemanticModel As SemanticModel, ByVal lTargetDocument As Document)
                mSemanticModel = lSemanticModel
                mTargetDocument = lTargetDocument
                mOptions = DirectCast(lTargetDocument?.Project.CompilationOptions, VisualBasicCompilationOptions)
            End Sub
            Public ReadOnly Property IsModule As Boolean
                Get
                    If IsModuleStack.Count = 0 Then
                        Return False
                    End If
                    Return IsModuleStack.Peek
                End Get
            End Property
            Private Shared Sub AddRefComment(node As CS.CSharpSyntaxNode, UnsupportedFeature As String)
                Dim StatementWithIssue As CS.CSharpSyntaxNode = GetStatementwithIssues(node)
                Dim LeadingTrivia As New List(Of SyntaxTrivia) From {
                    SyntaxFactory.CommentTrivia($"' TODO: VB does not support {UnsupportedFeature}."),
                    VB_EOLTrivia}
                Dim LeadingCommentStatement As VisualBasicSyntaxNode = SyntaxFactory.EmptyStatement.WithLeadingTrivia(LeadingTrivia)
                StatementWithIssue.AddMarker(LeadingCommentStatement, RemoveStatement:=False, AllowDuplicates:=True)
            End Sub

            Public Shared Function CommentOutUnsupportedStatements(node As CS.CSharpSyntaxNode, UnsupportedFeature As String) As Syntax.EmptyStatementSyntax
                Dim emptyStatementSyntax1 As Syntax.EmptyStatementSyntax = SyntaxFactory.EmptyStatement().WithConvertedLeadingTriviaFrom(node)
                Dim NewTrivia As New List(Of SyntaxTrivia)
                NewTrivia.AddRange(emptyStatementSyntax1.GetLeadingTrivia)
                NewTrivia.Add(VB_EOLTrivia)
                NewTrivia.Add(SyntaxFactory.CommentTrivia($"' TODO: VB does not support {UnsupportedFeature}."))
                NewTrivia.Add(VB_EOLTrivia)
                NewTrivia.Add(SyntaxFactory.CommentTrivia($"' Original Statement:"))
                NewTrivia.Add(VB_EOLTrivia)
                Dim NodeSplit() As String = Split(node.ToString, vbLf)
                For i As Integer = 0 To NodeSplit.Count - 1
                    NewTrivia.Add(SyntaxFactory.CommentTrivia($"' {NodeSplit(i)}"))
                    NewTrivia.Add(VB_EOLTrivia)
                Next
                emptyStatementSyntax1 = emptyStatementSyntax1.WithLeadingTrivia(NewTrivia).WithTrailingTrivia(VB_EOLTrivia)
                Return emptyStatementSyntax1
            End Function
            Public Overrides Function DefaultVisit(ByVal node As SyntaxNode) As VisualBasicSyntaxNode
                Throw New NotImplementedException(node.[GetType]().ToString & " not implemented!")
            End Function
            Public Overrides Function VisitCompilationUnit(ByVal node As CSS.CompilationUnitSyntax) As VisualBasicSyntaxNode
                For Each [using] As CSS.UsingDirectiveSyntax In node.Usings
                    [using].Accept(Me)
                Next
                Dim externList As New List(Of VisualBasicSyntaxNode)
                ' externlist is potentially a list of empty lines with trivia
                For Each extern As CSS.ExternAliasDirectiveSyntax In node.Externs
                    externList.Add(extern.Accept(Me))
                Next

                Dim attributes As SyntaxList(Of AttributesStatementSyntax) = SyntaxFactory.List(node.AttributeLists.Select(Function(a As CSS.AttributeListSyntax) SyntaxFactory.AttributesStatement(SyntaxFactory.SingletonList(DirectCast(a.Accept(Me), AttributeListSyntax)))))
                Dim members As SyntaxList(Of StatementSyntax) = SyntaxFactory.List(node.Members.Select(Function(m As CSS.MemberDeclarationSyntax) DirectCast(m.Accept(Me), StatementSyntax)))
                Dim compilationUnitSyntax1 As Syntax.CompilationUnitSyntax
                Dim EndOfFIleToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.EndOfFileToken).WithConvertedTriviaFrom(node.EndOfFileToken)

                If externList.Count > 0 Then
                    compilationUnitSyntax1 = SyntaxFactory.CompilationUnit(
                        SyntaxFactory.List(Of OptionStatementSyntax)(),
                        SyntaxFactory.List(allImports),
                        attributes,
                        members).WithTriviaFrom(externList(0))
                ElseIf allImports.Count > 0 Then
                    If members.Count > 0 AndAlso members(0).HasLeadingTrivia Then
                        If TypeOf members(0) IsNot NamespaceBlockSyntax OrElse members(0).GetLeadingTrivia.ToFullString.Contains("auto-generated") Then
                            Dim HeadingTriviaList As New List(Of SyntaxTrivia)
                            HeadingTriviaList.AddRange(members(0).GetLeadingTrivia)
                            If HeadingTriviaList(0).IsKind(VisualBasic.SyntaxKind.EndOfLineTrivia) Then
                                HeadingTriviaList.RemoveAt(0)
                                If HeadingTriviaList.Count > 0 Then
                                    HeadingTriviaList.Add(VB_EOLTrivia)
                                End If
                            End If
                            Dim NewMemberList As New SyntaxList(Of StatementSyntax)
                            NewMemberList = NewMemberList.Add(members(0).WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                            members = NewMemberList.AddRange(members.RemoveAt(0))
                            Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
                            ' Remove Leading whitespace
                            For Each t As SyntaxTrivia In HeadingTriviaList
                                If Not t.IsWhitespaceOrEndOfLine Then
                                    NewLeadingTrivia.Add(t)
                                End If
                            Next
                            allImports(0) = allImports(0).WithPrependedLeadingTrivia(NewLeadingTrivia)
                        End If
                    End If
                    compilationUnitSyntax1 = SyntaxFactory.CompilationUnit(
                    options:=SyntaxFactory.List(Of OptionStatementSyntax)(),
                    [imports]:=SyntaxFactory.List(allImports),
                    attributes:=attributes,
                    members:=members, endOfFileToken:=EndOfFIleToken)
                Else
                    compilationUnitSyntax1 = SyntaxFactory.CompilationUnit(
                        options:=SyntaxFactory.List(Of OptionStatementSyntax)(),
                        [imports]:=SyntaxFactory.List(allImports),
                        attributes:=attributes,
                        members:=members, endOfFileToken:=EndOfFIleToken)
                End If
                If MarkerError() Then
                    ' There are statements that were left out of translation
                    Throw New ApplicationException(GetMarkerErrorMessage)
                End If
                Return compilationUnitSyntax1
            End Function

            Public Overrides Function VisitDeclarationPattern(node As DeclarationPatternSyntax) As VisualBasicSyntaxNode
                ' Correct
                'INSTANT VB TODO TASK: VB has no equivalent to C# pattern variables in 'is' expressions:
                'ORIGINAL LINE: if (ModelExtensions.GetSymbolInfo(semanticModel, node).Symbol is IMethodSymbol methodSymbol && methodSymbol.ReturnType.Equals(ModelExtensions.GetTypeInfo(semanticModel, node).ConvertedType))
                'If TypeOf ModelExtensions.GetSymbolInfo(mSemanticModel, node).Symbol Is IMethodSymbol Then
                '    Dim methodSymbol As IMethodSymbol = DirectCast(ModelExtensions.GetSymbolInfo(mSemanticModel, node).Symbol, IMethodSymbol)
                '    If methodSymbol.ReturnType.Equals(ModelExtensions.GetTypeInfo(mSemanticModel, node).ConvertedType) Then
                '        Return SyntaxFactory.InvocationExpression(MemberAccessExpressionSyntax, SyntaxFactory.ArgumentList())
                '    End If
                'End If

                ' Current
                ' Dim methodSymbol As IMethodSymbol = DirectCast(ModelExtensions.GetSymbolInfo(mSemanticModel, node).Symbol, IMethodSymbol)

                Dim LeadingTrivia As New SyntaxTriviaList
                Dim StatementWithIssue As CS.CSharpSyntaxNode = GetStatementwithIssues(node)
                LeadingTrivia = CheckCorrectnessLeadingTrivia(StatementWithIssue, "VB has no direct equivalent To C# pattern variables 'is' expressions")
                Dim Designation As SingleVariableDesignationSyntax = DirectCast(node.Designation, SingleVariableDesignationSyntax)

                Dim value As ExpressionSyntax = SyntaxFactory.ParseExpression($"TryCast({node.Designation.Accept(Me).NormalizeWhitespace.ToFullString}, {node.Type.Accept(Me).NormalizeWhitespace.ToFullString})")

                Dim SeparatedSyntaxListOfModifiedIdentifier As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) =
                        SyntaxFactory.SingletonSeparatedList(
                            SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(Designation.Identifier, IsQualifiedName:=False, IsTypeName:=False)))
                Dim VariableType As Syntax.TypeSyntax = DirectCast(node.Type.Accept(Me), Syntax.TypeSyntax)

                Dim SeparatedListOfvariableDeclarations As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) =
                        SyntaxFactory.SingletonSeparatedList(
                            SyntaxFactory.VariableDeclarator(
                                SeparatedSyntaxListOfModifiedIdentifier,
                                SyntaxFactory.SimpleAsClause(VariableType),
                                SyntaxFactory.EqualsValue(NothingExpression)
                                    )
                             )

                Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(
                    SyntaxFactory.Token(SyntaxKind.DimKeyword))

                Dim DeclarationToBeAdded As Syntax.LocalDeclarationStatementSyntax =
                    SyntaxFactory.LocalDeclarationStatement(
                                        ModifiersTokenList,
                                        SeparatedListOfvariableDeclarations
                                        ).WithAdditionalAnnotations(Simplifier.Annotation).WithLeadingTrivia(LeadingTrivia)

                StatementWithIssue.AddMarker(DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=True)
                Return value
            End Function

            Public Overrides Function VisitImplicitElementAccess(ByVal node As CSS.ImplicitElementAccessSyntax) As VisualBasicSyntaxNode
                If node.ArgumentList.Arguments.Count > 1 Then
                    Throw New NotSupportedException("ImplicitElementAccess can only have one argument!")
                End If
                Return node.ArgumentList.Arguments(0).Expression.Accept(Me).WithConvertedTriviaFrom(node.ArgumentList.Arguments(0).Expression)
            End Function

            Public Overrides Function VisitLocalFunctionStatement(node As LocalFunctionStatementSyntax) As VisualBasicSyntaxNode
                Return MyBase.VisitLocalFunctionStatement(node)
            End Function
            Public Overrides Function VisitRefExpression(node As RefExpressionSyntax) As VisualBasicSyntaxNode
                AddRefComment(node, "RefExpressions")
                Return node.Expression.Accept(Me)
            End Function

            Public Overrides Function VisitRefType(node As RefTypeSyntax) As VisualBasicSyntaxNode
                AddRefComment(node, "RefType")
                Return node.Type.Accept(Me)
            End Function

            Public Overrides Function VisitRefTypeExpression(node As RefTypeExpressionSyntax) As VisualBasicSyntaxNode
                AddRefComment(node, "RefTypeExpressions")
                Return MyBase.VisitRefTypeExpression(node)
            End Function
            ''' <summary>
            '''
            ''' </summary>
            ''' <param name="Node"></param>
            ''' <returns></returns>
            ''' <remarks>Added by PC</remarks>
            Public Overrides Function VisitSingleVariableDesignation(ByVal Node As SingleVariableDesignationSyntax) As VisualBasicSyntaxNode
                Dim Identifier As SyntaxToken = GenerateSafeVBToken(Node.Identifier, IsQualifiedName:=False)
                Dim IdentifierExpression As Syntax.IdentifierNameSyntax = SyntaxFactory.IdentifierName(Identifier)
                Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(Identifier)
                Dim SeparatedSyntaxListOfModifiedIdentifier As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) =
                    SyntaxFactory.SingletonSeparatedList(
                        ModifiedIdentifier
                        )

                If Node.Parent.IsKind(CS.SyntaxKind.DeclarationExpression) Then
                    Dim Parent As DeclarationExpressionSyntax = DirectCast(Node.Parent, DeclarationExpressionSyntax)
                    Dim TypeName As VisualBasicSyntaxNode = Parent.Type.Accept(Me)

                    Dim SeparatedListOfvariableDeclarations As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) =
                        SyntaxFactory.SingletonSeparatedList(
                            SyntaxFactory.VariableDeclarator(
                                SeparatedSyntaxListOfModifiedIdentifier,
                                SyntaxFactory.SimpleAsClause(ConvertToType(TypeName.NormalizeWhitespace.ToString)),
                                SyntaxFactory.EqualsValue(NothingExpression)
                                    )
                             )

                    Dim ModifiersTokenList As SyntaxTokenList = SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.DimKeyword))
                    Dim DeclarationToBeAdded As Syntax.LocalDeclarationStatementSyntax =
                        SyntaxFactory.LocalDeclarationStatement(
                            ModifiersTokenList,
                            SeparatedListOfvariableDeclarations).WithAdditionalAnnotations(Simplifier.Annotation)

                    GetStatementwithIssues(Node).AddMarker(DeclarationToBeAdded, RemoveStatement:=False, AllowDuplicates:=True)
                ElseIf Node.Parent.IsKind(CS.SyntaxKind.DeclarationPattern) Then
                    Dim DeclarationPattern As DeclarationPatternSyntax = DirectCast(Node.Parent, DeclarationPatternSyntax)
                    Dim CasePatternSwitchLabel As CSS.SwitchLabelSyntax = DirectCast(DeclarationPattern.Parent, SwitchLabelSyntax)
                    Dim SwitchSection As CSS.SwitchSectionSyntax = DirectCast(CasePatternSwitchLabel.Parent, SwitchSectionSyntax)
                    Dim SwitchStatement As CSS.SwitchStatementSyntax = DirectCast(SwitchSection.Parent, CSS.SwitchStatementSyntax)
                    Dim SwitchExpression As ExpressionSyntax = DirectCast(SwitchStatement.Expression.Accept(Me), ExpressionSyntax)

                    Dim TypeName As Syntax.TypeSyntax = DirectCast(DeclarationPattern.Type.Accept(Me), Syntax.TypeSyntax)
                    Return SyntaxFactory.TypeOfIsExpression(SwitchExpression, TypeName).WithTrailingTrivia(SyntaxFactory.CommentTrivia($"' TODO: VB does not support Declaration Pattern. An attempt was made to convert the original Case Clause was ""case {Node.Parent.ToFullString}:"""))
                End If

                Return IdentifierExpression
            End Function
        End Class
    End Class
End Namespace
