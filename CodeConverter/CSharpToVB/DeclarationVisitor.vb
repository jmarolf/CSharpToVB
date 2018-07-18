Option Explicit On
Option Infer Off
Option Strict On
Imports System.Runtime.InteropServices
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports ArgumentListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArgumentListSyntax
Imports AttributeListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.AttributeListSyntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports IdentifierNameSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.IdentifierNameSyntax
Imports ParameterListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ParameterListSyntax
Imports ParameterSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ParameterSyntax
Imports StatementSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.StatementSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxFacts = Microsoft.CodeAnalysis.VisualBasic.SyntaxFacts
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeParameterListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeParameterListSyntax
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax
Imports VisualBasicExtensions = Microsoft.CodeAnalysis.VisualBasic.VisualBasicExtensions

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)
            Private Const CompilerServices As String = "System.Runtime.CompilerServices"

            Private Shared Function ConvertOperatorDeclarationToken(ByVal lSyntaxKind As CS.SyntaxKind) As SyntaxKind
                Select Case lSyntaxKind
                    Case CS.SyntaxKind.AmpersandToken
                        Return SyntaxKind.AndKeyword
                    Case CS.SyntaxKind.BarToken
                        Return SyntaxKind.OrKeyword
                    Case CS.SyntaxKind.EqualsEqualsToken
                        Return SyntaxKind.EqualsToken
                    Case CS.SyntaxKind.EqualsGreaterThanToken
                        Return SyntaxKind.GreaterThanEqualsToken
                    Case CS.SyntaxKind.ExclamationEqualsToken
                        Return SyntaxKind.LessThanGreaterThanToken
                    Case CS.SyntaxKind.GreaterThanEqualsToken
                        Return SyntaxKind.GreaterThanEqualsToken
                    Case CS.SyntaxKind.GreaterThanToken
                        Return SyntaxKind.GreaterThanToken
                    Case CS.SyntaxKind.LessThanToken
                        Return SyntaxKind.LessThanToken
                    Case CS.SyntaxKind.LessThanEqualsToken
                        Return SyntaxKind.LessThanEqualsToken
                    Case CS.SyntaxKind.MinusToken
                        Return SyntaxKind.MinusToken
                    Case CS.SyntaxKind.PlusToken
                        Return SyntaxKind.PlusToken
                End Select
                Throw New NotSupportedException()
            End Function

            Private Shared Function GetSemicolonTrivia(semicolonToken As SyntaxToken) As List(Of SyntaxTrivia)
                Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                If semicolonToken.HasLeadingTrivia Then
                    NewTrailingTrivia.AddRange(ExtractComments(ConvertTrivia(semicolonToken.LeadingTrivia), Leading:=True))
                End If
                If semicolonToken.HasTrailingTrivia Then
                    NewTrailingTrivia.AddRange(ExtractComments(ConvertTrivia(semicolonToken.TrailingTrivia), Leading:=False))
                End If
                Return NewTrailingTrivia
            End Function

            Private Shared Function SpecialReservedWord(v As String) As Boolean

                If v.Equals("Alias", StringComparison.InvariantCultureIgnoreCase) OrElse
                   v.Equals("End", StringComparison.InvariantCultureIgnoreCase) OrElse
                   v.Equals("Imports", StringComparison.InvariantCultureIgnoreCase) OrElse
                   v.Equals("Module", StringComparison.InvariantCultureIgnoreCase) OrElse
                   v.Equals("Option", StringComparison.InvariantCultureIgnoreCase) OrElse
                   v.Equals("Optional", StringComparison.InvariantCultureIgnoreCase) Then
                    Return True
                End If
                Return False
            End Function
            Private Function ConvertAccessor(ByVal node As CSS.AccessorDeclarationSyntax, IsModule As Boolean, ByRef isIterator As Boolean) As AccessorBlockSyntax
                Dim blockKind As SyntaxKind
                Dim stmt As AccessorStatementSyntax
                Dim endStmt As EndBlockStatementSyntax
                Dim body As SyntaxList(Of StatementSyntax) = SyntaxFactory.List(Of StatementSyntax)()
                isIterator = False
                If node.Body IsNot Nothing Then
                    Dim visitor As MethodBodyVisitor = New MethodBodyVisitor(mSemanticModel, Me)
                    body = SyntaxFactory.List(node.Body.Statements.SelectMany(Function(s As CSS.StatementSyntax) s.Accept(visitor)))
                    isIterator = visitor.IsInterator
                End If
                Dim attributes As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(node.AttributeLists.Select(Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax)))
                Dim modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.Local)
                Dim parent As CSS.BasePropertyDeclarationSyntax = DirectCast(node.Parent.Parent, CSS.BasePropertyDeclarationSyntax)
                Dim valueParam As ParameterSyntax

                Select Case CS.CSharpExtensions.Kind(node)
                    Case CS.SyntaxKind.GetAccessorDeclaration
                        blockKind = SyntaxKind.GetAccessorBlock
                        stmt = SyntaxFactory.GetAccessorStatement(attributes, modifiers, parameterList:=Nothing)
                        endStmt = SyntaxFactory.EndGetStatement()
                    Case CS.SyntaxKind.SetAccessorDeclaration
                        blockKind = SyntaxKind.SetAccessorBlock
                        valueParam = SyntaxFactory.Parameter(SyntaxFactory.ModifiedIdentifier("value")).
                            WithAsClause(SyntaxFactory.SimpleAsClause(DirectCast(parent.Type.Accept(Me), TypeSyntax))).
                            WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.ByValKeyword)))
                        stmt = SyntaxFactory.SetAccessorStatement(attributes, modifiers, SyntaxFactory.ParameterList(SyntaxFactory.SingletonSeparatedList(valueParam)))
                        endStmt = SyntaxFactory.EndSetStatement()
                    Case CS.SyntaxKind.AddAccessorDeclaration
                        blockKind = SyntaxKind.AddHandlerAccessorBlock
                        valueParam = SyntaxFactory.Parameter(SyntaxFactory.ModifiedIdentifier("value")).
                            WithAsClause(SyntaxFactory.SimpleAsClause(DirectCast(parent.Type.Accept(Me), TypeSyntax))).
                            WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.ByValKeyword)))
                        stmt = SyntaxFactory.AddHandlerAccessorStatement(attributes, modifiers, SyntaxFactory.ParameterList(SyntaxFactory.SingletonSeparatedList(valueParam)))
                        endStmt = SyntaxFactory.EndAddHandlerStatement()
                    Case CS.SyntaxKind.RemoveAccessorDeclaration
                        blockKind = SyntaxKind.RemoveHandlerAccessorBlock
                        valueParam = SyntaxFactory.Parameter(SyntaxFactory.ModifiedIdentifier("value")).
                            WithAsClause(SyntaxFactory.SimpleAsClause(DirectCast(parent.Type.Accept(Me), TypeSyntax))).
                            WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.ByValKeyword)))
                        stmt = SyntaxFactory.RemoveHandlerAccessorStatement(attributes, modifiers, SyntaxFactory.ParameterList(SyntaxFactory.SingletonSeparatedList(valueParam)))
                        endStmt = SyntaxFactory.EndRemoveHandlerStatement()
                    Case Else
                        Throw New NotSupportedException()
                End Select
                Return SyntaxFactory.AccessorBlock(blockKind, stmt, body, endStmt).WithConvertedTriviaFrom(node)
            End Function

            Private Sub ConvertAndSplitAttributes(ByVal attributeLists As SyntaxList(Of CSS.AttributeListSyntax), <Out> ByRef attributes As SyntaxList(Of AttributeListSyntax), <Out> ByRef returnAttributes As SyntaxList(Of AttributeListSyntax))
                Dim retAttr As List(Of AttributeListSyntax) = New List(Of AttributeListSyntax)()
                Dim attr As List(Of AttributeListSyntax) = New List(Of AttributeListSyntax)()
                For Each attrList As CSS.AttributeListSyntax In attributeLists
                    If attrList.Target IsNot Nothing AndAlso SyntaxTokenExtensions.IsKind(attrList.Target.Identifier, CS.SyntaxKind.ReturnKeyword) = True Then
                        ' Remove trailing CRLF from return attributes
                        retAttr.Add(DirectCast(attrList.Accept(Me).With({SyntaxFactory.Space}, {SyntaxFactory.Space}), AttributeListSyntax))
                    Else
                        attr.Add(DirectCast(attrList.Accept(Me), AttributeListSyntax))
                    End If
                Next

                returnAttributes = SyntaxFactory.List(retAttr)
                attributes = SyntaxFactory.List(attr)
            End Sub
            Public Overrides Function VisitAnonymousObjectMemberDeclarator(ByVal node As CSS.AnonymousObjectMemberDeclaratorSyntax) As VisualBasicSyntaxNode
#Disable Warning CC0013 ' Use Ternary operator.
                If node.NameEquals Is Nothing Then
#Enable Warning CC0013 ' Use Ternary operator.
                    Return SyntaxFactory.InferredFieldInitializer(DirectCast(node.Expression.Accept(Me), ExpressionSyntax)).WithConvertedTriviaFrom(node)
                Else
                    Return SyntaxFactory.NamedFieldInitializer(keyKeyword:=SyntaxFactory.Token(SyntaxKind.KeyKeyword),
                                                               dotToken:=SyntaxFactory.Token(SyntaxKind.DotToken),
                                                               name:=DirectCast(node.NameEquals.Name.Accept(Me), IdentifierNameSyntax),
                                                               equalsToken:=SyntaxFactory.Token(SyntaxKind.EqualsToken),
                                                               expression:=DirectCast(node.Expression.Accept(Me), ExpressionSyntax)
                                                               ).WithConvertedTriviaFrom(node)
                End If
            End Function

            Public Overrides Function VisitArrowExpressionClause(node As ArrowExpressionClauseSyntax) As VisualBasicSyntaxNode
                Return node.Expression.Accept(Me)
            End Function

            Public Overrides Function VisitConstructorDeclaration(ByVal node As CSS.ConstructorDeclarationSyntax) As VisualBasicSyntaxNode
                Dim AttributeList As SyntaxList(Of AttributeListSyntax) = SyntaxFactory.List(
                        node.AttributeLists.[Select](Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax)))
                Dim Initializer As Syntax.ExpressionStatementSyntax = Nothing
                If node.Initializer IsNot Nothing Then
                    Initializer = DirectCast(node.Initializer.Accept(Me), Syntax.ExpressionStatementSyntax)
                End If
                Dim modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.[New])

                Dim parameterList As ParameterListSyntax = DirectCast(node.ParameterList?.Accept(Me), ParameterListSyntax)
                Dim SubNewStatement As SubNewStatementSyntax = SyntaxFactory.SubNewStatement(
                                                                    AttributeList, modifiers, parameterList)

                Dim Body As New SyntaxList(Of StatementSyntax)
                Dim CloseBraceToken As SyntaxToken = CS.SyntaxFactory.Token(CSharp.SyntaxKind.CloseBraceToken)
                If node.Body IsNot Nothing Then
                    Body = SyntaxFactory.List(node.Body.Statements.SelectMany(Function(s As CSS.StatementSyntax) s.Accept(New MethodBodyVisitor(mSemanticModel, Me))))
                    CloseBraceToken = node.Body.CloseBraceToken
                End If
                If Initializer IsNot Nothing Then
                    Body = Body.InsertRange(0, ReplaceStatementWithMarkedStatement(node:=node,
                                                                                   Statements:=SyntaxFactory.SingletonList(Of StatementSyntax)(Initializer)))
                End If
                Dim EndSubStatement As EndBlockStatementSyntax = SyntaxFactory.EndSubStatement().WithConvertedTriviaFrom(CloseBraceToken)
                Return SyntaxFactory.ConstructorBlock(SubNewStatement, Body, EndSubStatement).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitConstructorInitializer(node As ConstructorInitializerSyntax) As VisualBasicSyntaxNode
                Dim ArgumentList As ArgumentListSyntax = DirectCast((node.ArgumentList.Accept(Me)), ArgumentListSyntax)
                Dim SimpleMemberAccessExpression As Syntax.MemberAccessExpressionSyntax
#Disable Warning CC0014 ' Use Ternary operator.
                If TypeOf node.Parent.Parent Is CSS.StructDeclarationSyntax Then
#Enable Warning CC0014 ' Use Ternary operator.
                    SimpleMemberAccessExpression = SyntaxFactory.SimpleMemberAccessExpression(SyntaxFactory.MeExpression(), SyntaxFactory.IdentifierName("New"))
                Else
                    SimpleMemberAccessExpression = SyntaxFactory.SimpleMemberAccessExpression(SyntaxFactory.MyBaseExpression(), SyntaxFactory.IdentifierName("New"))
                End If

                Dim ExpressionStatement As Syntax.ExpressionStatementSyntax = SyntaxFactory.ExpressionStatement(SyntaxFactory.InvocationExpression(SimpleMemberAccessExpression, ArgumentList))
                Return ExpressionStatement
            End Function

            ''' <summary>
            ''' Creates a new object initialized to a meaningful value.
            ''' </summary>
            ''' <param name="value"></param>
            'INSTANT VB TODO TASK: Generic operators are not available in VB:
            'ORIGINAL LINE: public static implicit operator @Optional<T>(T value)
            'Public Shared Widening Operator DirectCast(value As T) As [Optional]
            '    Return New [Optional](Of T)(value)
            'End Operator
            Public Overrides Function VisitConversionOperatorDeclaration(ByVal node As CSS.ConversionOperatorDeclarationSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.EmptyStatement.WithLeadingTrivia(node.ConvertNodeToMultipleLineComment("TODO TASK: Generic operators are not available in VB, original lines are below, no translation was attempted")).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitDestructorDeclaration(ByVal node As CSS.DestructorDeclarationSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.SubBlock(SyntaxFactory.SubStatement(
                                              SyntaxFactory.List(node.AttributeLists.Select(Function(a As CSS.AttributeListSyntax) DirectCast(a.Accept(Me), AttributeListSyntax))),
                                              SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.ProtectedKeyword),
                                                                      SyntaxFactory.Token(SyntaxKind.OverridesKeyword)
                                                                      ),
                                              SyntaxFactory.Identifier(NameOf(Finalize)),
                                              Nothing,
                                              DirectCast(node.ParameterList?.Accept(Me), ParameterListSyntax),
                                              Nothing,
                                              Nothing,
                                              Nothing),
                                              SyntaxFactory.List(node.Body.Statements.SelectMany(Function(s As CSS.StatementSyntax) s.Accept(New MethodBodyVisitor(mSemanticModel, Me))))
                                              ).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitEventDeclaration(ByVal node As CSS.EventDeclarationSyntax) As VisualBasicSyntaxNode
                Dim attributes As SyntaxList(Of AttributeListSyntax) = Nothing
                Dim returnAttributes As SyntaxList(Of AttributeListSyntax) = Nothing
                ConvertAndSplitAttributes(node.AttributeLists, attributes, returnAttributes)
                Dim Modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.Member)
                Dim Identifier As SyntaxToken = GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False)
                Dim AsClause As SimpleAsClauseSyntax = SyntaxFactory.SimpleAsClause(attributeLists:=returnAttributes, type:=DirectCast(node.Type.Accept(Me), TypeSyntax))
                Modifiers = Modifiers.Add(SyntaxFactory.Token(SyntaxKind.CustomKeyword))
                Dim stmt As EventStatementSyntax = SyntaxFactory.EventStatement(attributeLists:=attributes, modifiers:=Modifiers, identifier:=Identifier, parameterList:=Nothing, asClause:=AsClause, implementsClause:=Nothing)
                If node.AccessorList.Accessors.All(Function(a As CSS.AccessorDeclarationSyntax) a.Body Is Nothing) Then
                    Return stmt.WithConvertedTriviaFrom(node)
                End If
                Dim accessors As AccessorBlockSyntax() = node.AccessorList?.Accessors.[Select](Function(a As CSS.AccessorDeclarationSyntax) ConvertAccessor(a, IsModule:=IsModule, isIterator:=False)).ToArray()
                Return SyntaxFactory.EventBlock(stmt, SyntaxFactory.List(accessors)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitEventFieldDeclaration(ByVal node As CSS.EventFieldDeclarationSyntax) As VisualBasicSyntaxNode
                Dim decl As CSS.VariableDeclaratorSyntax = node.Declaration.Variables.Single()
                Dim id As SyntaxToken = SyntaxFactory.Identifier(decl.Identifier.ValueText, SyntaxFacts.IsKeywordKind(VisualBasicExtensions.Kind(decl.Identifier)), decl.Identifier.GetIdentifierText(), TypeCharacter.None)
                Dim attributes As SyntaxList(Of AttributeListSyntax) = Nothing
                Dim returnAttributes As SyntaxList(Of AttributeListSyntax) = Nothing
                ConvertAndSplitAttributes(node.AttributeLists, attributes, returnAttributes)
                Return SyntaxFactory.EventStatement(attributes, ConvertModifiers(node.Modifiers, IsModule, TokenContext.Member), id, parameterList:=Nothing, SyntaxFactory.SimpleAsClause(returnAttributes, DirectCast(node.Declaration.Type.Accept(Me), TypeSyntax)), implementsClause:=Nothing).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitFieldDeclaration(ByVal node As CSS.FieldDeclarationSyntax) As VisualBasicSyntaxNode
                Dim _TypeInfo As TypeInfo = ModelExtensions.GetTypeInfo(mSemanticModel, node.Declaration.Type)
                Dim variableOrConstOrReadonly As TokenContext = TokenContext.VariableOrConst
                If _TypeInfo.ConvertedType IsNot Nothing AndAlso _TypeInfo.ConvertedType.TypeKind = TypeKind.Class Then
                    For i As Integer = 0 To node.Declaration.Variables.Count - 1
                        Dim v As CSS.VariableDeclaratorSyntax = node.Declaration.Variables(i)
                        If v.Initializer IsNot Nothing AndAlso v.Initializer.Value.Kind = CS.SyntaxKind.NullLiteralExpression Then
                            variableOrConstOrReadonly = TokenContext.Readonly
                        End If
                    Next
                End If
                Dim modifierList As New List(Of SyntaxToken)
                modifierList.AddRange(ConvertModifiers(node.Modifiers, IsModule, variableOrConstOrReadonly))
                If modifierList.Count = 0 Then
                    modifierList.Add(SyntaxFactory.Token(SyntaxKind.PrivateKeyword))
                End If
                Dim LeadingTrivia As New List(Of SyntaxTrivia)
                If node.Modifiers.Count > 0 Then
                    LeadingTrivia.AddRange(ConvertTrivia(node.Modifiers(0).LeadingTrivia))
                End If
                Dim modifiers As SyntaxTokenList = SyntaxFactory.TokenList(modifierList)
                Dim AttributeLists As New SyntaxList(Of AttributeListSyntax)
                For Each a As CSS.AttributeListSyntax In node.AttributeLists
                    AttributeLists = AttributeLists.Add(DirectCast(a.Accept(Me), AttributeListSyntax))
                Next
                Dim declarators As SeparatedSyntaxList(Of Syntax.VariableDeclaratorSyntax) = RemodelVariableDeclaration(node.Declaration, Me, LeadingTrivia)
                Dim FieldDeclaration As Syntax.FieldDeclarationSyntax
                If node.Modifiers.Contains(Function(t As SyntaxToken) t.IsKind(CS.SyntaxKind.VolatileKeyword)) Then
                    Dim Name As TypeSyntax = SyntaxFactory.ParseTypeName("Volatile")
                    Dim VolatileAttribute As SeparatedSyntaxList(Of Syntax.AttributeSyntax) = SyntaxFactory.SingletonSeparatedList(SyntaxFactory.Attribute(Name))
                    LeadingTrivia.Add(SyntaxFactory.CommentTrivia("' TODO TASK: VB has no direct equivalent to C# Volatile"))
                    AttributeLists = AttributeLists.Add(SyntaxFactory.AttributeList(VolatileAttribute).WithLeadingTrivia(LeadingTrivia))
                End If
                FieldDeclaration = SyntaxFactory.FieldDeclaration(AttributeLists, modifiers, declarators).With(LeadingTrivia, GetSemicolonTrivia(node.SemicolonToken))
                Return AddSpecialCommentToField(node, FieldDeclaration)
            End Function
            Public Overrides Function VisitIndexerDeclaration(ByVal node As CSS.IndexerDeclarationSyntax) As VisualBasicSyntaxNode
                Dim id As SyntaxToken = SyntaxFactory.Identifier("Item")
                Dim attributes As SyntaxList(Of AttributeListSyntax) = Nothing
                Dim returnAttributes As SyntaxList(Of AttributeListSyntax) = Nothing
                ConvertAndSplitAttributes(node.AttributeLists, attributes, returnAttributes)
                Dim isIterator As Boolean = False
                Dim accessors As List(Of AccessorBlockSyntax) = New List(Of AccessorBlockSyntax)()
                If node.AccessorList IsNot Nothing Then
                    For Each a As CSS.AccessorDeclarationSyntax In node.AccessorList.Accessors
                        Dim _isIterator As Boolean
                        accessors.Add(ConvertAccessor(a, IsModule, _isIterator))
                        isIterator = isIterator Or _isIterator
                    Next
                End If

                Dim modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, TokenContext.Member).Insert(0, SyntaxFactory.Token(SyntaxKind.DefaultKeyword))
                Select Case accessors.Count
                    Case 0
                        Dim AccessorStatement As AccessorStatementSyntax = SyntaxFactory.GetAccessorStatement()
                        Dim Expression As ExpressionSyntax = DirectCast(node.ExpressionBody.Accept(Me), ExpressionSyntax)
                        Dim Body As SyntaxList(Of StatementSyntax) = SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.ReturnStatement(Expression))
                        Dim EndStmt As EndBlockStatementSyntax = SyntaxFactory.EndGetStatement()
                        accessors.Add(SyntaxFactory.AccessorBlock(SyntaxKind.GetAccessorBlock, AccessorStatement, Body, EndStmt))
                    Case 1
                        Dim NeedKeyword As Boolean = True
                        If accessors(0).AccessorStatement.Kind() = SyntaxKind.GetAccessorStatement Then
                            For Each keyword As SyntaxToken In modifiers
                                If keyword.ValueText = "ReadOnly" Then
                                    NeedKeyword = False
                                    Exit For
                                End If
                            Next
                            If NeedKeyword Then
                                modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.ReadOnlyKeyword))
                            End If
                        Else
                            For Each keyword As SyntaxToken In modifiers
                                If keyword.ValueText = "WriteOnly" Then
                                    NeedKeyword = False
                                    Exit For
                                End If
                            Next
                            If NeedKeyword Then
                                modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.WriteOnlyKeyword))
                            End If
                        End If
                    Case 2
                    Case Else
                        Throw UnreachableException
                End Select
                If isIterator Then
                    modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.IteratorKeyword))
                End If
                Dim parameterList As ParameterListSyntax = DirectCast(node.ParameterList?.Accept(Me), ParameterListSyntax).WithRemovedTrailingEOLTrivia
                Dim AsClause As SimpleAsClauseSyntax = SyntaxFactory.SimpleAsClause(returnAttributes, DirectCast(node.Type.Accept(Me), TypeSyntax))
                Dim stmt As PropertyStatementSyntax = SyntaxFactory.PropertyStatement(attributes, modifiers, id, parameterList, AsClause, initializer:=Nothing, implementsClause:=Nothing)
                Dim accessorList As AccessorListSyntax = node.AccessorList
                If accessorList Is Nothing OrElse accessorList.Accessors.All(Function(a As CSS.AccessorDeclarationSyntax) a.Body Is Nothing) Then
                    If node.ExpressionBody Is Nothing Then
                        Return stmt
                    End If
                End If
                Return SyntaxFactory.PropertyBlock(stmt, SyntaxFactory.List(accessors))
            End Function

            Public Overrides Function VisitMethodDeclaration(node As CSS.MethodDeclarationSyntax) As VisualBasicSyntaxNode
                If node.Modifiers.Any(Function(m As SyntaxToken) SyntaxTokenExtensions.IsKind(m, CS.SyntaxKind.UnsafeKeyword)) Then
                    Return CommentOutUnsupportedStatements(node, "Unchecked Functions")
                End If
                UsedIdentifierStack.Push(UsedIdentifiers)
                Dim id As SyntaxToken = GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False)
                Dim visitor As New MethodBodyVisitor(mSemanticModel, Me)

                Dim methodInfo As ISymbol = ModelExtensions.GetDeclaredSymbol(mSemanticModel, node)
                Dim ReturnVoid As Boolean? = methodInfo?.GetReturnType()?.SpecialType = SpecialType.System_Void
                Dim containingType As INamedTypeSymbol = methodInfo?.ContainingType
                Dim attributes As SyntaxList(Of AttributeListSyntax) = Nothing
                Dim returnAttributes As SyntaxList(Of AttributeListSyntax) = Nothing
                ConvertAndSplitAttributes(node.AttributeLists, attributes, returnAttributes)
                Dim parameterList As ParameterListSyntax = DirectCast(node.ParameterList?.Accept(Me), ParameterListSyntax)

                Dim block As SyntaxList(Of StatementSyntax)? = Nothing

                If node.Body IsNot Nothing Then
                    Dim StatementNodes As New List(Of StatementSyntax)
                    For Each s As CSS.StatementSyntax In node.Body.Statements
                        Dim StatementCollection As SyntaxList(Of StatementSyntax) = s.Accept(visitor)
                        StatementNodes.AddRange(StatementCollection)
                    Next
                    block = SyntaxFactory.List(StatementNodes)
                ElseIf node.ExpressionBody IsNot Nothing Then
                    Dim Expression1 As VisualBasicSyntaxNode = node.ExpressionBody.Expression.Accept(Me)
                    Dim syntaxList1 As SyntaxList(Of StatementSyntax)
                    If ReturnVoid Then
                        Select Case Expression1.Kind
                            Case SyntaxKind.SimpleAssignmentStatement
                                syntaxList1 = SyntaxFactory.SingletonList(DirectCast(Expression1, StatementSyntax))
                            Case SyntaxKind.InvocationExpression
                                syntaxList1 = SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.CallStatement(DirectCast(Expression1, ExpressionSyntax)))
                            Case SyntaxKind.ConditionalAccessExpression
                                syntaxList1 = SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.ExpressionStatement(DirectCast(Expression1, ExpressionSyntax)))
                            Case Else
                                syntaxList1 = Nothing
                                Stop
                        End Select
                    Else
                        syntaxList1 = SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.ReturnStatement(DirectCast(Expression1, ExpressionSyntax)))
                    End If
                    block = ReplaceStatementWithMarkedStatement(node, syntaxList1)
                End If
                If node.Modifiers.Any(Function(m As SyntaxToken) SyntaxTokenExtensions.IsKind(m, CS.SyntaxKind.ExternKeyword)) Then
                    block = SyntaxFactory.List(Of StatementSyntax)()
                End If
                Dim modifiers As SyntaxTokenList = ConvertModifiers(node.Modifiers, IsModule, If(containingType?.IsInterfaceType() = True, TokenContext.Local, TokenContext.Member))
                If visitor.IsInterator Then
                    modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.IteratorKeyword))
                End If
                If node.ParameterList.Parameters.Count > 0 AndAlso node.ParameterList.Parameters(0).Modifiers.Any(CS.SyntaxKind.ThisKeyword) Then
                    attributes = attributes.Insert(0, SyntaxFactory.AttributeList(SyntaxFactory.SingletonSeparatedList(SyntaxFactory.Attribute(Nothing, SyntaxFactory.ParseTypeName("Extension"), SyntaxFactory.ArgumentList()))))
                    If Not DirectCast(node.SyntaxTree, CS.CSharpSyntaxTree).HasUsingDirective(CompilerServices) Then
                        Dim ImportComilierServices As ImportsStatementSyntax = SyntaxFactory.ImportsStatement(SyntaxFactory.SingletonSeparatedList(Of ImportsClauseSyntax)(SyntaxFactory.SimpleImportsClause(SyntaxFactory.ParseName(CompilerServices)))).WithAppendedTrailingTrivia(VB_EOLTrivia)
                        If Not allImports.ContainsName(CompilerServices) Then
                            allImports.Add(ImportComilierServices)
                        End If
                    End If
                End If
                If containingType?.IsStatic = True Then
                    modifiers = SyntaxFactory.TokenList(modifiers.Where(Function(t As SyntaxToken) Not (t.IsKind(SyntaxKind.SharedKeyword, SyntaxKind.PublicKeyword))))
                End If

                Dim FunctionLeadingTrivia As New List(Of SyntaxTrivia)
                FunctionLeadingTrivia.AddRange(ConvertTrivia(node.GetLeadingTrivia))
                Dim FunctionTrailingTrivia As New List(Of SyntaxTrivia)
                If node.Body Is Nothing Then
                    FunctionTrailingTrivia.AddRange(ConvertTrivia(node.GetTrailingTrivia))
                Else
                    If node.Body.OpenBraceToken.HasTrailingTrivia Then
                        FunctionTrailingTrivia.AddRange(ConvertTrivia(node.Body.OpenBraceToken.TrailingTrivia))
                    Else
                        FunctionTrailingTrivia.AddRange(ConvertTrivia(node.GetTrailingTrivia))
                    End If
                End If

                Dim EndSubOrFunctionStatement As EndBlockStatementSyntax
                Dim stmt As MethodStatementSyntax
                If ReturnVoid Then
                    EndSubOrFunctionStatement = SyntaxFactory.EndSubStatement()
                    If node.Body IsNot Nothing Then
                        EndSubOrFunctionStatement = EndSubOrFunctionStatement.WithConvertedTriviaFrom(node.Body.CloseBraceToken)
                    End If
                    Dim TypeParameterList1 As TypeParameterListSyntax = DirectCast(node.TypeParameterList?.Accept(Me), TypeParameterListSyntax)
                    stmt = SyntaxFactory.SubStatement(
                        attributes,
                        modifiers,
                        id,
                        TypeParameterList1,
                        parameterList,
                        asClause:=Nothing,
                        handlesClause:=Nothing,
                        implementsClause:=Nothing).With(FunctionLeadingTrivia, FunctionTrailingTrivia)
                    UsedIdentifiers = DirectCast(UsedIdentifierStack.Pop, Dictionary(Of String, SymbolTableEntry))
                    If block Is Nothing Then
                        If modifiers.Contains(Function(t As SyntaxToken) t.IsKind(SyntaxKind.PartialKeyword)) Then
                            Return SyntaxFactory.SubBlock(stmt, statements:=Nothing, endSubOrFunctionStatement:=EndSubOrFunctionStatement).With(FunctionLeadingTrivia, FunctionTrailingTrivia)
                        End If
                        Return stmt.With(FunctionLeadingTrivia, FunctionTrailingTrivia)
                    End If
                    Return SyntaxFactory.SubBlock(stmt, block.Value, EndSubOrFunctionStatement).With(FunctionLeadingTrivia, FunctionTrailingTrivia)
                End If
                EndSubOrFunctionStatement = SyntaxFactory.EndFunctionStatement()
                If node.Body IsNot Nothing Then
                    EndSubOrFunctionStatement = EndSubOrFunctionStatement.WithConvertedTriviaFrom(node.Body.CloseBraceToken)
                End If
                Dim type As TypeSyntax = DirectCast(node.ReturnType.Accept(Me), TypeSyntax)
                If type.ToString.StartsWith("[") Then
                    Dim S As String() = type.ToString.Split({"["c, "]"}, StringSplitOptions.RemoveEmptyEntries)
                    If Not SpecialReservedWord(S(0)) Then
                        type = SyntaxFactory.ParseTypeName(type.ToString.Replace("]", "").Replace("[", ""))
                    End If
                End If
                Dim AsClause As SimpleAsClauseSyntax = SyntaxFactory.SimpleAsClause(returnAttributes, type.WithLeadingTrivia(SyntaxFactory.Space))
                Dim ParamListTrailingTrivia As SyntaxTriviaList = parameterList.GetTrailingTrivia
                parameterList = parameterList.WithTrailingTrivia(AsClause.GetTrailingTrivia)
                AsClause = AsClause.WithTrailingTrivia(ParamListTrailingTrivia)
                Dim TypeParameterList As TypeParameterListSyntax = DirectCast(node.TypeParameterList?.Accept(Me), TypeParameterListSyntax)

                stmt = SyntaxFactory.FunctionStatement(
                                            attributes,
                                            modifiers,
                                            id,
                                            TypeParameterList,
                                            parameterList,
                                            AsClause,
                                            handlesClause:=Nothing,
                                            implementsClause:=Nothing).With(FunctionLeadingTrivia, FunctionTrailingTrivia)
                UsedIdentifiers = DirectCast(UsedIdentifierStack.Pop, Dictionary(Of String, SymbolTableEntry))
                If block Is Nothing Then
                    Return stmt.With(FunctionLeadingTrivia, FunctionTrailingTrivia)
                End If
                Return SyntaxFactory.FunctionBlock(stmt, block.Value, EndSubOrFunctionStatement).With(FunctionLeadingTrivia, FunctionTrailingTrivia)
            End Function

            Public Overrides Function VisitOperatorDeclaration(ByVal node As CSS.OperatorDeclarationSyntax) As VisualBasicSyntaxNode
                Dim attributes As SyntaxList(Of AttributeListSyntax) = Nothing
                Dim returnAttributes As SyntaxList(Of AttributeListSyntax) = Nothing
                ConvertAndSplitAttributes(node.AttributeLists, attributes, returnAttributes)
                Dim visitor As New MethodBodyVisitor(mSemanticModel, Me)
                Dim body As New SyntaxList(Of StatementSyntax)
                If node.Body IsNot Nothing Then
                    body = SyntaxFactory.List(node.Body.Statements.SelectMany(Function(s As CSS.StatementSyntax) s.Accept(visitor)))
                ElseIf node.ExpressionBody IsNot Nothing Then
                    body = SyntaxFactory.SingletonList(Of StatementSyntax)(SyntaxFactory.ReturnStatement(DirectCast(node.ExpressionBody.Accept(Me), ExpressionSyntax)))
                Else
                    Stop
                End If
                Dim parameterList As ParameterListSyntax = DirectCast(node.ParameterList?.Accept(Me), ParameterListSyntax).WithRemovedTrailingEOLTrivia
                Dim stmt As OperatorStatementSyntax = SyntaxFactory.OperatorStatement(attributes, ConvertModifiers(node.Modifiers, IsModule, TokenContext.Member), SyntaxFactory.Token(ConvertOperatorDeclarationToken(CS.CSharpExtensions.Kind(node.OperatorToken))), parameterList, SyntaxFactory.SimpleAsClause(returnAttributes, DirectCast(node.ReturnType.Accept(Me), TypeSyntax)))
                Return SyntaxFactory.OperatorBlock(stmt, body).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitPropertyDeclaration(ByVal node As CSS.PropertyDeclarationSyntax) As VisualBasicSyntaxNode
                Dim ImplementsClause As ImplementsClauseSyntax = Nothing
                Dim id As SyntaxToken
                Dim ExplicateInterfaceIdentifier As Syntax.QualifiedNameSyntax = Nothing
                If node.ExplicitInterfaceSpecifier IsNot Nothing Then
                    Dim VisualBasicSyntaxNode1 As VisualBasicSyntaxNode = node.ExplicitInterfaceSpecifier.Accept(Me)
                    If TypeOf VisualBasicSyntaxNode1 Is Syntax.QualifiedNameSyntax Then
                        ExplicateInterfaceIdentifier = DirectCast(VisualBasicSyntaxNode1, Syntax.QualifiedNameSyntax)
                    Else
                        ' Might need to handle other types here (Generic Name, Identifier..)
                    End If
                End If
                ' It might be possible to combined the following if statement with the one above but for now keep them separate
                If ExplicateInterfaceIdentifier Is Nothing Then
                    id = GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False)
                Else
                    id = SyntaxFactory.Identifier($"{ExplicateInterfaceIdentifier.Right.ToString}_{node.Identifier.ValueText}")
                    Dim SimpleName As IdentifierNameSyntax = SyntaxFactory.IdentifierName(node.Identifier.ValueText)
                    Dim QualifiedName As Syntax.QualifiedNameSyntax = SyntaxFactory.QualifiedName(ExplicateInterfaceIdentifier, SimpleName)
                    Dim interfaceMembers As SeparatedSyntaxList(Of Syntax.QualifiedNameSyntax)
                    interfaceMembers = interfaceMembers.Add(QualifiedName)
                    ImplementsClause = SyntaxFactory.ImplementsClause(interfaceMembers)
                End If
                Dim attributes As SyntaxList(Of AttributeListSyntax) = Nothing
                Dim returnAttributes As SyntaxList(Of AttributeListSyntax) = Nothing
                ConvertAndSplitAttributes(node.AttributeLists, attributes, returnAttributes)
                Dim isIterator As Boolean = False
                Dim accessors As List(Of AccessorBlockSyntax) = New List(Of AccessorBlockSyntax)()
                If node.ExpressionBody IsNot Nothing Then
                    Dim ReturnedExpression As ExpressionSyntax = DirectCast(node.ExpressionBody.Expression.Accept(Me), ExpressionSyntax).WithConvertedTriviaFrom(node.ExpressionBody)
                    Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                    Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
                    Dim ReturnStatement As Syntax.ReturnStatementSyntax = SyntaxFactory.ReturnStatement(ReturnedExpression.WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                    NewLeadingTrivia.AddRange(ReturnStatement.GetLeadingTrivia)
                    For Each t As SyntaxTrivia In ReturnStatement.GetTrailingTrivia
                        Select Case t.Kind
                            Case SyntaxKind.WhitespaceTrivia, SyntaxKind.EndOfLineTrivia
                                NewTrailingTrivia.Add(t)
                            Case SyntaxKind.IfDirectiveTrivia
                                NewLeadingTrivia.Add(t)
                            Case Else
                                Stop
                        End Select
                    Next
                    ReturnStatement = ReturnStatement.With(NewLeadingTrivia, NewTrailingTrivia)
                    Dim Statements As SyntaxList(Of StatementSyntax) = ReplaceStatementWithMarkedStatement(node, SyntaxFactory.SingletonList(Of StatementSyntax)(ReturnStatement))
                    accessors.Add(SyntaxFactory.AccessorBlock(SyntaxKind.GetAccessorBlock, SyntaxFactory.GetAccessorStatement(), Statements, SyntaxFactory.EndGetStatement()))
                Else
                    If node.AccessorList IsNot Nothing Then
                        For Each a As CSS.AccessorDeclarationSyntax In node.AccessorList.Accessors
                            Dim _isIterator As Boolean
                            accessors.Add(ConvertAccessor(a, IsModule, _isIterator))
                            isIterator = isIterator Or _isIterator
                        Next
                    End If
                End If

                Dim CSharpModifiers As SyntaxTokenList = node.Modifiers
                If node.AccessorList IsNot Nothing AndAlso node.AccessorList.Accessors.Count = 1 Then
                    Select Case node.AccessorList.Accessors(0).Keyword.ToString
                        Case "get"
                            CSharpModifiers = CSharpModifiers.Add(CS.SyntaxFactory.Token(CS.SyntaxKind.ReadOnlyKeyword))
                        Case "set"
                        Case Else
                            Throw UnreachableException
                    End Select

                End If
                Dim Context As TokenContext = TokenContext.Member
                ' TODO find better way to find out if we are in interface
                If node.IsParentKind(CS.SyntaxKind.InterfaceDeclaration) Then
                    Context = TokenContext.InterfaceOrModule
                End If
                Dim modifiers As SyntaxTokenList = ConvertModifiers(CSharpModifiers, IsModule, Context)
                If isIterator Then
                    modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.IteratorKeyword))
                End If

                If node.AccessorList Is Nothing Then
                    modifiers = modifiers.Add(SyntaxFactory.Token(SyntaxKind.ReadOnlyKeyword))
                End If

                Dim stmt As PropertyStatementSyntax = SyntaxFactory.PropertyStatement(attributeLists:=attributes,
                                                                                      modifiers:=modifiers,
                                                                                      identifier:=id,
                                                                                      parameterList:=Nothing,
                                                                                      asClause:=SyntaxFactory.SimpleAsClause(returnAttributes, DirectCast(node.Type.Accept(Me).WithoutTrivia, TypeSyntax)),
                                                                                      initializer:=If(node.Initializer Is Nothing, Nothing, SyntaxFactory.EqualsValue(DirectCast(node.Initializer.Value.Accept(Me), ExpressionSyntax))),
                                                                                      implementsClause:=ImplementsClause
                                                                                      )
                Dim StmtList As SyntaxList(Of StatementSyntax) = ReplaceStatementWithMarkedStatement(node, SyntaxFactory.SingletonList(Of StatementSyntax)(stmt))
                Dim AddedLeadingTrivia As New List(Of SyntaxTrivia)
                If StmtList.Count = 2 Then
                    AddedLeadingTrivia.AddRange(StmtList(0).GetLeadingTrivia)
                End If
                If node.AccessorList IsNot Nothing AndAlso node.AccessorList.Accessors.All(Function(a As CSS.AccessorDeclarationSyntax) a.Body Is Nothing) = True Then
                    Return stmt.WithConvertedTriviaFrom(node).WithPrependedLeadingTrivia(AddedLeadingTrivia)
                End If
                Return SyntaxFactory.PropertyBlock(stmt, SyntaxFactory.List(accessors)).WithConvertedTriviaFrom(node).WithPrependedLeadingTrivia
            End Function
        End Class
    End Class
End Namespace
