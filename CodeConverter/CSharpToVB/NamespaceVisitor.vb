Option Explicit On
Option Infer Off
Option Strict On

Imports IVisualBasicCode.CodeConverter.Util

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Imports AttributeListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.AttributeListSyntax
Imports CS = Microsoft.CodeAnalysis.CSharp

Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax

Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports NameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.NameSyntax
Imports ParameterListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ParameterListSyntax
Imports StatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.StatementSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeParameterListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeParameterListSyntax
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)

            Private Const InlineAssignHelperCode As String = "<Obsolete(""Please refactor code that uses this function, it is a simple work-around to simulate inline assignment in VB!"")>
Private Shared Function __InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
target = value
Return value
End Function"

            Private Sub ConvertBaseList(ByVal type As CSS.BaseTypeDeclarationSyntax, ByVal [inherits] As List(Of InheritsStatementSyntax), ByVal [implements] As List(Of ImplementsStatementSyntax))
                Dim arr As TypeSyntax()
                Select Case type.Kind()
                    Case CS.SyntaxKind.ClassDeclaration
                        Dim classOrInterface As CSS.TypeSyntax = type.BaseList?.Types.FirstOrDefault()?.Type
                        If classOrInterface Is Nothing Then
                            Return
                        End If
                        Dim classOrInterfaceSymbol As ISymbol = ModelExtensions.GetSymbolInfo(mSemanticModel, classOrInterface).Symbol
                        If classOrInterfaceSymbol?.IsInterfaceType() Then
                            arr = type.BaseList?.Types.[Select](Function(t As BaseTypeSyntax) DirectCast(t.Type.Accept(Me), TypeSyntax)).ToArray()
                            If arr.Length > 0 Then [implements].Add(SyntaxFactory.ImplementsStatement(arr))
                        Else
                            [inherits].Add(SyntaxFactory.InheritsStatement(DirectCast(classOrInterface.Accept(Me), TypeSyntax)))
                            arr = type.BaseList?.Types.Skip(1).[Select](Function(t As BaseTypeSyntax) DirectCast(t.Type.Accept(Me), TypeSyntax)).ToArray()
                            If arr.Length > 0 Then [implements].Add(SyntaxFactory.ImplementsStatement(arr))
                        End If

                    Case CS.SyntaxKind.StructDeclaration
                        arr = type.BaseList?.Types.[Select](Function(t As BaseTypeSyntax) DirectCast(t.Type.Accept(Me), TypeSyntax)).ToArray()
                        If arr?.Length > 0 Then [implements].Add(SyntaxFactory.ImplementsStatement(arr))
                    Case CS.SyntaxKind.InterfaceDeclaration
                        arr = type.BaseList?.Types.[Select](Function(t As BaseTypeSyntax) DirectCast(t.Type.Accept(Me), TypeSyntax)).ToArray()
                        If arr?.Length > 0 Then [inherits].Add(SyntaxFactory.InheritsStatement(arr))
                End Select
            End Sub

            Private Iterator Function PatchInlineHelpers(ByVal node As CSS.BaseTypeDeclarationSyntax, IsModule As Boolean) As IEnumerable(Of StatementSyntax)
                If inlineAssignHelperMarkers.Contains(node) Then
                    inlineAssignHelperMarkers.Remove(node)
                    Yield TryCast(SyntaxFactory.ParseSyntaxTree(InlineAssignHelperCode.Replace("Shared ", If(IsModule, "", "Shared "))).GetRoot().ChildNodes().FirstOrDefault(), StatementSyntax)
                End If
            End Function

            Public Overrides Function VisitClassDeclaration(ByVal node As CSS.ClassDeclarationSyntax) As VisualBasicSyntaxNode
                Dim saveUsedIdentifiers As Dictionary(Of String, SymbolTableEntry) = UsedIdentifiers
                UsedIdentifierStack.Push(UsedIdentifiers)
                IsModuleStack.Push(node.Modifiers.Any(CS.SyntaxKind.StaticKeyword) And node.TypeParameterList Is Nothing)
                Dim members As List(Of StatementSyntax) = node.Members.Select(Function(m As CSS.MemberDeclarationSyntax) DirectCast(m.Accept(Me), StatementSyntax)).ToList()
                Dim id As SyntaxToken = GenerateSafeVBToken(node.Identifier, False)

                Dim [inherits] As List(Of InheritsStatementSyntax) = New List(Of InheritsStatementSyntax)()
                Dim [implements] As List(Of ImplementsStatementSyntax) = New List(Of ImplementsStatementSyntax)()

                ConvertBaseList(node, [inherits], [implements])
                members.AddRange(PatchInlineHelpers(node, IsModule))

                Dim attributeLists As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax)))
                Dim typeParameterList As TypeParameterListSyntax = DirectCast(node.TypeParameterList?.Accept(Me), TypeParameterListSyntax)

                If IsModule Then
                    Dim EndModule As EndBlockStatementSyntax = SyntaxFactory.EndModuleStatement().WithConvertedTriviaFrom(node.CloseBraceToken)
                    Dim ModuleModifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.InterfaceOrModule)
                    Dim ModuleBlock As ModuleBlockSyntax = SyntaxFactory.ModuleBlock(
                        SyntaxFactory.ModuleStatement(attributeLists, ModuleModifiers, id, typeParameterList),
                        SyntaxFactory.List([inherits]),
                        SyntaxFactory.List([implements]),
                        SyntaxFactory.List(members),
                        EndModule).WithConvertedTriviaFrom(node)
                    UsedIdentifiers = saveUsedIdentifiers
                    IsModuleStack.Pop()
                    UsedIdentifiers = DirectCast(UsedIdentifierStack.Pop, Dictionary(Of String, SymbolTableEntry))
                    Return ModuleBlock
                Else
                    Dim EndClass As EndBlockStatementSyntax = SyntaxFactory.EndClassStatement().WithConvertedTriviaFrom(node.CloseBraceToken)
                    Dim ClassModifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.Class)
                    Dim ClassBlock As ClassBlockSyntax = SyntaxFactory.ClassBlock(
                        SyntaxFactory.ClassStatement(
                            attributeLists,
                            ClassModifiers,
                            id,
                            typeParameterList),
                            SyntaxFactory.List([inherits]),
                            SyntaxFactory.List([implements]),
                            SyntaxFactory.List(members), EndClass).WithConvertedTriviaFrom(node)
                    UsedIdentifiers = saveUsedIdentifiers
                    IsModuleStack.Pop()
                    UsedIdentifiers = DirectCast(UsedIdentifierStack.Pop, Dictionary(Of String, SymbolTableEntry))
                    Return ClassBlock
                End If
            End Function

            Public Overrides Function VisitDelegateDeclaration(ByVal node As CSS.DelegateDeclarationSyntax) As VisualBasicSyntaxNode
                Dim id As SyntaxToken = GenerateSafeVBToken(node.Identifier, False)
                Dim methodInfo As INamedTypeSymbol = TryCast(ModelExtensions.GetDeclaredSymbol(mSemanticModel, node), INamedTypeSymbol)
                Dim AttributeLists As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax)))
                Dim Modifiers As SyntaxTokenList = ConvertModifiers(modifiers:=node.Modifiers, IsModule:=IsModule)
                Dim TypeParameterList As TypeParameterListSyntax = DirectCast(node.TypeParameterList?.Accept(Me), TypeParameterListSyntax)
                Dim ParameterList As ParameterListSyntax = DirectCast(node.ParameterList?.Accept(Me), ParameterListSyntax).WithoutTrailingTrivia
                If methodInfo.DelegateInvokeMethod.GetReturnType()?.SpecialType = SpecialType.System_Void Then
                    Return SyntaxFactory.DelegateSubStatement(attributeLists:=AttributeLists, modifiers:=Modifiers, identifier:=id, typeParameterList:=TypeParameterList, parameterList:=ParameterList, asClause:=Nothing).WithConvertedTriviaFrom(node)
                Else
                    Dim AsClause As SimpleAsClauseSyntax = SyntaxFactory.SimpleAsClause(DirectCast(node.ReturnType.Accept(Me), TypeSyntax).WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                    Return SyntaxFactory.DelegateFunctionStatement(attributeLists:=AttributeLists, modifiers:=Modifiers, identifier:=id, typeParameterList:=TypeParameterList, parameterList:=ParameterList, asClause:=AsClause).WithConvertedTriviaFrom(node)
                End If
            End Function

            Public Overrides Function VisitEnumDeclaration(ByVal node As CSS.EnumDeclarationSyntax) As VisualBasicSyntaxNode
                Dim members As New List(Of StatementSyntax)
                Dim CS_Members As SeparatedSyntaxList(Of CSS.EnumMemberDeclarationSyntax) = node.Members
                If CS_Members.Count > 0 Then

                    Dim CS_Separators As New List(Of SyntaxToken)
                    CS_Separators.AddRange(node.Members.GetSeparators)
                    For i As Integer = 0 To CS_Members.Count - 2
                        members.Add(DirectCast(CS_Members(i).Accept(Me).WithConvertedTrailingTriviaFrom(CS_Separators(i)), Syntax.StatementSyntax))
                    Next
                    members.Add(DirectCast(CS_Members.Last.Accept(Me).WithConvertedTrailingTriviaFrom(CS_Members.Last), Syntax.StatementSyntax))
                End If

                Dim baseType As TypeSyntax = DirectCast(node.BaseList?.Types.Single().Accept(Me), TypeSyntax)
                Return SyntaxFactory.EnumBlock(SyntaxFactory.EnumStatement(SyntaxFactory.List(node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax))), ConvertModifiers(node.Modifiers, IsModule), GenerateSafeVBToken(node.Identifier, False), If(baseType Is Nothing, Nothing, SyntaxFactory.SimpleAsClause(baseType))), SyntaxFactory.List(members)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitEnumMemberDeclaration(ByVal node As CSS.EnumMemberDeclarationSyntax) As VisualBasicSyntaxNode
                Dim initializer As ExpressionSyntax = DirectCast(node.EqualsValue?.Value.Accept(Me), ExpressionSyntax)
                Return SyntaxFactory.EnumMemberDeclaration(SyntaxFactory.List(node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax))), GenerateSafeVBToken(node.Identifier, False), initializer:=If(initializer Is Nothing, Nothing, SyntaxFactory.EqualsValue(initializer))).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitExplicitInterfaceSpecifier(node As ExplicitInterfaceSpecifierSyntax) As VisualBasicSyntaxNode
                Return node.Name.Accept(Me)
            End Function

            Public Overrides Function VisitExternAliasDirective(node As ExternAliasDirectiveSyntax) As VisualBasicSyntaxNode
                Dim emptyStatementSyntax1 As Syntax.EmptyStatementSyntax = CommentOutUnsupportedStatements(node, "Extern Alias")
                Return emptyStatementSyntax1
            End Function

            Public Overrides Function VisitInterfaceDeclaration(ByVal node As CSS.InterfaceDeclarationSyntax) As VisualBasicSyntaxNode
                Dim [inherits] As List(Of InheritsStatementSyntax) = New List(Of InheritsStatementSyntax)()
                Dim [implements] As List(Of ImplementsStatementSyntax) = New List(Of ImplementsStatementSyntax)()
                ConvertBaseList(node, [inherits], [implements])
                Dim AttributeList As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax)))
                Dim Modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.InterfaceOrModule)
                If node.Modifiers.ToString.Contains("unsafe") Then
                    Return CommentOutUnsupportedStatements(node, "unsafe interfaces")
                End If
                Dim members As StatementSyntax() = node.Members.[Select](Function(m As CSS.MemberDeclarationSyntax) DirectCast(m.Accept(Me), StatementSyntax)).ToArray()
                Dim Identifier As SyntaxToken = GenerateSafeVBToken(node.Identifier, False)
                Dim TypeParameterList As TypeParameterListSyntax = DirectCast(node.TypeParameterList?.Accept(Me), TypeParameterListSyntax)
                Dim InterfaceStatement As InterfaceStatementSyntax = SyntaxFactory.InterfaceStatement(AttributeList, Modifiers, Identifier, TypeParameterList)
                Dim InheritsStatementList As SyntaxList(Of InheritsStatementSyntax) = SyntaxFactory.List([inherits])
                Dim ImplementsStatementList As SyntaxList(Of ImplementsStatementSyntax) = SyntaxFactory.List([implements])
                Dim StatementList As SyntaxList(Of StatementSyntax) = SyntaxFactory.List(members)
                Dim EndInterfaceStatement As EndBlockStatementSyntax = SyntaxFactory.EndInterfaceStatement.WithConvertedLeadingTriviaFrom(node.CloseBraceToken)
                Return SyntaxFactory.InterfaceBlock(InterfaceStatement,
                                                    InheritsStatementList,
                                                    ImplementsStatementList,
                                                    StatementList,
                                                    EndInterfaceStatement
                                                    ).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitNamespaceDeclaration(ByVal node As CSS.NamespaceDeclarationSyntax) As VisualBasicSyntaxNode
                For Each [using] As CSS.UsingDirectiveSyntax In node.Usings
                    [using].Accept(Me)
                Next

                For Each extern As CSS.ExternAliasDirectiveSyntax In node.Externs
                    extern.Accept(Me)
                Next

                Dim members As IEnumerable(Of StatementSyntax) = node.Members.[Select](Function(m As CSS.MemberDeclarationSyntax) DirectCast(m.Accept(Me), StatementSyntax))
                Dim NamespaceKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.NamespaceKeyword)
                Dim Trivia As New List(Of SyntaxTrivia) From {
                    VB_EOLTrivia
                }
                Trivia.AddRange(ConvertTrivia(node.OpenBraceToken.LeadingTrivia))
                Trivia.AddRange(ConvertTrivia(node.OpenBraceToken.TrailingTrivia))
                Dim NamespaceStatement As NamespaceStatementSyntax = SyntaxFactory.NamespaceStatement(NamespaceKeyword, DirectCast(node.Name.Accept(Me), NameSyntax)).WithTrailingTrivia(Trivia)
                Dim members1 As SyntaxList(Of StatementSyntax) = SyntaxFactory.List(members)
                Dim EndNamespaceStatement As EndBlockStatementSyntax = SyntaxFactory.EndNamespaceStatement.WithConvertedTriviaFrom(node.CloseBraceToken)
                Dim namespaceBlock As NamespaceBlockSyntax = SyntaxFactory.NamespaceBlock(NamespaceStatement, members1, EndNamespaceStatement).WithConvertedTriviaFrom(node)
                Return namespaceBlock
            End Function

            Public Overrides Function VisitStructDeclaration(ByVal node As CSS.StructDeclarationSyntax) As VisualBasicSyntaxNode
                Dim members As List(Of StatementSyntax) = node.Members.[Select](Function(m As CSS.MemberDeclarationSyntax) DirectCast(m.Accept(Me), StatementSyntax)).ToList()
                Dim [inherits] As List(Of InheritsStatementSyntax) = New List(Of InheritsStatementSyntax)()
                Dim [implements] As List(Of ImplementsStatementSyntax) = New List(Of ImplementsStatementSyntax)()
                ConvertBaseList(node, [inherits], [implements])
                members.AddRange(PatchInlineHelpers(node, IsModule))
                If members.Any Then
                    members(0) = members(0).WithPrependedLeadingTrivia(ConvertTrivia(node.OpenBraceToken.LeadingTrivia))
                End If
                Dim AttributeLists As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax)))
                Dim TypeParameterList As TypeParameterListSyntax = DirectCast(node.TypeParameterList?.Accept(Me), TypeParameterListSyntax)
                Dim StructureBlock As StructureBlockSyntax = SyntaxFactory.StructureBlock(
                                                    SyntaxFactory.StructureStatement(AttributeLists,
                                                                                     ConvertModifiers(node.Modifiers, IsModule),
                                                                                     GenerateSafeVBToken(node.Identifier, False),
                                                                                     TypeParameterList
                                                                                    ),
                                                                    SyntaxFactory.List([inherits]),
                                                                    SyntaxFactory.List([implements]),
                                                                    SyntaxFactory.List(members),
                                                    SyntaxFactory.EndStructureStatement.WithConvertedTriviaFrom(node.CloseBraceToken)
                                                    ).WithConvertedTriviaFrom(node)
                If node.Modifiers.Contains(Function(t As SyntaxToken) t.IsKind(CS.SyntaxKind.UnsafeKeyword)) Then
                    StructureBlock = StructureBlock.WithPrependedLeadingTrivia(SyntaxFactory.CommentTrivia("' TODO TASK: VB has no direct equivalent to C# Unsafe Structure"))
                End If
                Return StructureBlock
            End Function

            Public Overrides Function VisitUsingDirective(ByVal node As CSS.UsingDirectiveSyntax) As VisualBasicSyntaxNode
                UsedIdentifierStack.Push(UsedIdentifiers)
                UsedIdentifiers.Clear()
                Dim [alias] As ImportAliasClauseSyntax = Nothing
                If node.[Alias] IsNot Nothing Then
                    Dim name As CSS.IdentifierNameSyntax = node.[Alias].Name
                    Dim id As SyntaxToken = GenerateSafeVBToken(name.Identifier, False)
                    [alias] = SyntaxFactory.ImportAliasClause(id)
                End If

                Dim clause As ImportsClauseSyntax = SyntaxFactory.SimpleImportsClause([alias], DirectCast(node.Name.Accept(Me), NameSyntax))

                Dim import As ImportsStatementSyntax = SyntaxFactory.ImportsStatement(SyntaxFactory.SingletonSeparatedList(clause)).WithConvertedTriviaFrom(node)

                allImports.Add(import)
                UsedIdentifiers = DirectCast(UsedIdentifierStack.Pop, Dictionary(Of String, SymbolTableEntry))
                Return Nothing
            End Function

        End Class

    End Class

End Namespace