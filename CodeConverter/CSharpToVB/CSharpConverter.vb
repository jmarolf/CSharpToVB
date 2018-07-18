Option Explicit On
Option Infer Off
Option Strict On

Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports ExpressionSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxFacts = Microsoft.CodeAnalysis.VisualBasic.SyntaxFacts
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax
Imports VariableDeclaratorSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.VariableDeclaratorSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Private Const UnicodeCloseQuote As Char = ChrW(&H201D)

        Private Const UnicodeDoubleCloseQuote As String = UnicodeCloseQuote & UnicodeCloseQuote

        Private Const UnicodeDoubleOpenQuote As String = UnicodeOpenQuote & UnicodeOpenQuote

        Private Const UnicodeOpenQuote As Char = ChrW(&H201C)

        Private Shared UsedIdentifierStack As New Stack(New Dictionary(Of String, SymbolTableEntry))

        Public Shared thisLock As New Object

        Public Shared UsedIdentifiers As New Dictionary(Of String, SymbolTableEntry)(StringComparer.Ordinal)

        Public Enum TokenContext
            [Global]
            InterfaceOrModule
            [Class]
            Member
            VariableOrConst
            Local
            [New]
            [Readonly]
            XMLComment
        End Enum
        Private Shared Function Binary(Value As Byte) As String
            Return $"&B{System.Convert.ToString(Value, 2).PadLeft(8, "0"c)}"
        End Function

        Private Shared Function Binary(Value As SByte) As String
            Return $"&B{System.Convert.ToString(Value, 2).PadLeft(8, "0"c)}"
        End Function

        Private Shared Function Binary(Value As Short) As String
            Return $"&B{System.Convert.ToString(Value, 2).PadLeft(16, "0"c)}"
        End Function

        Private Shared Function Binary(Value As UShort) As String
            Return $"&B{System.Convert.ToString(Value, 2).PadLeft(16, "0"c)}"
        End Function

        Private Shared Function Binary(Value As Integer) As String
            Return $"&B{System.Convert.ToString(Value, 2).PadLeft(32, "0"c)}"
        End Function

        Private Shared Function Binary(Value As UInteger) As String
            Return $"&B{System.Convert.ToString(Value, 2).PadLeft(32, "0"c)}UI"
        End Function

        Private Shared Function Binary(Value As Long) As String
            Return $"&B{System.Convert.ToString(Value, 2).PadLeft(64, "0"c)}"
        End Function

        Private Shared Function ConvertModifier(ByVal m As SyntaxToken, ByVal IsModule As Boolean, ByVal Optional context As TokenContext = TokenContext.[Global]) As SyntaxToken?
            Dim TokenKind As SyntaxKind = ConvertToken(CS.CSharpExtensions.Kind(m), IsModule, context)
            Return If(TokenKind = SyntaxKind.None, Nothing, New SyntaxToken?(SyntaxFactory.Token(TokenKind)))
        End Function

        Private Shared Function ConvertModifiers(ByVal modifiers As IEnumerable(Of SyntaxToken), ByVal IsModule As Boolean, ByVal Optional context As TokenContext = TokenContext.[Global]) As SyntaxTokenList
            Return SyntaxFactory.TokenList(ConvertModifiersCore(modifiers, IsModule, context))
        End Function

        Private Shared Function ConvertModifiers(ByVal modifiers As SyntaxTokenList, ByVal IsModule As Boolean, ByVal Optional context As TokenContext = TokenContext.[Global]) As SyntaxTokenList
            Return SyntaxFactory.TokenList(ConvertModifiersCore(modifiers, IsModule, context))
        End Function

        Private Shared Iterator Function ConvertModifiersCore(ByVal modifiers As IEnumerable(Of SyntaxToken), ByVal IsModule As Boolean, ByVal context As TokenContext) As IEnumerable(Of SyntaxToken)
            If context <> TokenContext.Local AndAlso context <> TokenContext.InterfaceOrModule AndAlso context <> TokenContext.Class Then
                Dim visibility As Boolean = False
                For Each token As SyntaxToken In modifiers
                    If IsVisibility(token, context) Then
                        visibility = True
                        Exit For
                    End If
                Next

                If Not visibility AndAlso context = TokenContext.Member Then
                    Yield CSharpDefaultVisibility(context)
                End If
            End If

            For Each token As SyntaxToken In modifiers.Where(Function(m As SyntaxToken) Not IgnoreInContext(m, context))
                Dim m As SyntaxToken? = ConvertModifier(token, IsModule, context)
                If m.HasValue Then
                    Yield m.Value
                End If
            Next

        End Function

        Private Shared Function CSharpDefaultVisibility(ByVal context As TokenContext) As SyntaxToken
            Select Case context
                Case TokenContext.[Global]
                    Return SyntaxFactory.Token(SyntaxKind.FriendKeyword)
                Case TokenContext.Local, TokenContext.Member, TokenContext.VariableOrConst
                    Return SyntaxFactory.Token(SyntaxKind.PrivateKeyword)
                Case TokenContext.[New]
                    Return SyntaxFactory.Token(SyntaxKind.EmptyToken)
            End Select

            Throw New ArgumentOutOfRangeException(NameOf(context))
        End Function

        Private Shared Function ExtractIdentifier(ByVal v As CSS.VariableDeclaratorSyntax) As ModifiedIdentifierSyntax
            Return SyntaxFactory.ModifiedIdentifier(GenerateSafeVBToken(v.Identifier, IsQualifiedName:=False))
        End Function

        ''' <summary>
        ''' Returns Safe VB Name
        ''' </summary>
        ''' <param name="id">Original Name</param>
        ''' <param name="IsQualifiedName">True if name is part of a Qualified Name and should not be renamed</param>
        ''' <param name="IsTypeName"></param>
        ''' <returns></returns>
        Private Shared Function GenerateSafeVBToken(ByVal id As SyntaxToken, IsQualifiedName As Boolean, Optional IsTypeName As Boolean = False) As SyntaxToken
            Dim keywordKind As SyntaxKind = SyntaxFacts.GetKeywordKind(id.ValueText)
            If SyntaxFacts.IsKeywordKind(keywordKind) Then
                Return MakeIdentifierUnique(id, BracketNeeded:=True, QualifiedNameOrTypeName:=IsQualifiedName)
            End If
            If Not IsTypeName Then
                If SyntaxFacts.IsPredefinedType(keywordKind) Then
                    Return MakeIdentifierUnique(id, BracketNeeded:=True, QualifiedNameOrTypeName:=IsQualifiedName)
                End If
            End If
            If id.IsParentKind(CSharp.SyntaxKind.Parameter) Then
                Dim MethodDeclaration As CSS.MethodDeclarationSyntax = TryCast(id.Parent.Parent?.Parent, CSS.MethodDeclarationSyntax)
                IsQualifiedName = If(MethodDeclaration IsNot Nothing AndAlso String.Compare(MethodDeclaration.Identifier.ValueText, id.ValueText, ignoreCase:=True) = 0, False, True)
            End If
            Return MakeIdentifierUnique(id, BracketNeeded:=False, QualifiedNameOrTypeName:=IsQualifiedName)
        End Function

        Private Shared Function GetLiteralExpression(ByVal value As Object, Token As SyntaxToken) As ExpressionSyntax
            If TypeOf value Is Boolean Then
                Return If(CBool(value), SyntaxFactory.TrueLiteralExpression(SyntaxFactory.Token(SyntaxKind.TrueKeyword)), SyntaxFactory.FalseLiteralExpression(SyntaxFactory.Token(SyntaxKind.FalseKeyword)))
            End If
            If Token.ToString.StartsWith("0x") Then
                Dim HEXValueString As String = $"&H{Token.ToString.Substring(2)}".Replace("ul", "", StringComparison.OrdinalIgnoreCase).Replace("u", "", StringComparison.OrdinalIgnoreCase)
                If TypeOf value Is SByte Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(HEXValueString, CSByte(value)))
                If TypeOf value Is Short Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(HEXValueString, CShort(value)))
                If TypeOf value Is UShort Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(HEXValueString, CUShort(value)))
                If TypeOf value Is Integer Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(HEXValueString, CInt(value)))
                If TypeOf value Is UInteger Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(HEXValueString & "UI", CUInt(value)))
                If TypeOf value Is Long Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(HEXValueString, CLng(value)))
                If TypeOf value Is ULong Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(HEXValueString & "UL", CULng(value)))
            ElseIf Token.ToString.StartsWith("0b") Then
                If TypeOf value Is SByte Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal($"{Binary(CSByte(value))}", CSByte(value)))
                If TypeOf value Is Short Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal($"{Binary(CShort(value))}", CShort(value)))
                If TypeOf value Is UShort Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal($"{Binary(CUShort(value))}", CUShort(value)))
                If TypeOf value Is Integer Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(($"{Binary(CInt(value))}"), CInt(value)))
                If TypeOf value Is UInteger Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal($"{Binary(CUInt(value))}", CUInt(value)))
                If TypeOf value Is Long Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal($"{Binary(CLng(value))}", CLng(value)))
                If TypeOf value Is ULong Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal($"{Binary(CType(CULng(value), Long))}UL", CULng(value)))
            End If

            If TypeOf value Is Byte Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CByte(value)))
            If TypeOf value Is SByte Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CSByte(value)))
            If TypeOf value Is Short Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CShort(value)))
            If TypeOf value Is UShort Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CUShort(value)))
            If TypeOf value Is Integer Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CInt(value)))
            If TypeOf value Is UInteger Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CUInt(value)))
            If TypeOf value Is Long Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CLng(value)))
            If TypeOf value Is ULong Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CULng(value)))
            If TypeOf value Is Single Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CSng(value)))
            If TypeOf value Is Double Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CDbl(value)))
            If TypeOf value Is Decimal Then Return SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(CDec(value)))
            If TypeOf value Is Char Then
                If AscW(CChar(value)) = &H201C Then
                    Return SyntaxFactory.LiteralExpression(SyntaxKind.CharacterLiteralExpression, SyntaxFactory.Literal($"{UnicodeOpenQuote}{UnicodeOpenQuote}"))
                End If
                If AscW(CChar(value)) = &H201D Then
                    Return SyntaxFactory.LiteralExpression(SyntaxKind.CharacterLiteralExpression, SyntaxFactory.Literal($"{UnicodeCloseQuote}{UnicodeCloseQuote}"))
                End If
                Return SyntaxFactory.LiteralExpression(SyntaxKind.CharacterLiteralExpression, SyntaxFactory.Literal(CChar(value)))
            End If
            If TypeOf value Is String Then
                Dim StrValue As String = DirectCast(value, String)
                If StrValue.Contains("\") Then
                    StrValue = ConvertCSharpEscapes(StrValue)
                End If
                If StrValue.Contains("{") Or StrValue.Contains("}") Then
                    StrValue = StrValue
                End If
                If StrValue.Contains(UnicodeOpenQuote) Then
                    StrValue = StrValue.ConverUnicodeQuotes(UnicodeOpenQuote)
                End If
                If StrValue.Contains(UnicodeCloseQuote) Then
                    StrValue = StrValue.ConverUnicodeQuotes(UnicodeCloseQuote)
                End If
                Return SyntaxFactory.LiteralExpression(SyntaxKind.StringLiteralExpression, SyntaxFactory.Literal(StrValue))
            End If
            If value Is Nothing Then
                Return SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(SyntaxKind.NothingKeyword))
            End If
            Return Nothing
        End Function

        Private Shared Function IgnoreInContext(ByVal m As SyntaxToken, ByVal context As TokenContext) As Boolean
            Select Case context
                Case TokenContext.InterfaceOrModule
                    Return m.IsKind(CS.SyntaxKind.PublicKeyword, CS.SyntaxKind.StaticKeyword)
                Case TokenContext.Class
                    Return m.IsKind(CS.SyntaxKind.StaticKeyword)
            End Select

            Return False
        End Function

        Private Shared Function IsVisibility(ByVal token As SyntaxToken, ByVal context As TokenContext) As Boolean
            Return token.IsKind(CS.SyntaxKind.PublicKeyword, CS.SyntaxKind.InternalKeyword, CS.SyntaxKind.ProtectedKeyword, CS.SyntaxKind.PrivateKeyword) OrElse
                (context = TokenContext.VariableOrConst AndAlso SyntaxTokenExtensions.IsKind(token, CS.SyntaxKind.ConstKeyword))
        End Function

        Private Shared Function Literal(ByVal value As Object) As ExpressionSyntax
            Return GetLiteralExpression(value, New SyntaxToken)
        End Function

        Private Shared Function MakeIdentifierUnique(id As SyntaxToken, BracketNeeded As Boolean, QualifiedNameOrTypeName As Boolean) As SyntaxToken
            Dim ConvertedIdentifier As String = If(BracketNeeded, $"[{id.ValueText}]", id.ValueText)
            If ConvertedIdentifier = "_" Then
                ConvertedIdentifier = "underscore"
            End If
            ' Don't Change Qualified Names
            If QualifiedNameOrTypeName Then
                If Not UsedIdentifiers.ContainsKey(ConvertedIdentifier) Then
                    UsedIdentifiers.Add(ConvertedIdentifier, New SymbolTableEntry(ConvertedIdentifier, True))
                End If
                Return SyntaxFactory.Identifier(UsedIdentifiers(ConvertedIdentifier).Name)
            End If
            If UsedIdentifiers.ContainsKey(ConvertedIdentifier) Then
                ' We have a case sensitive exact match so just return it
                Return SyntaxFactory.Identifier(UsedIdentifiers(ConvertedIdentifier).Name)
            End If
            For Each ident As KeyValuePair(Of String, SymbolTableEntry) In UsedIdentifiers
                If String.Compare(ident.Key, ConvertedIdentifier, ignoreCase:=False) = 0 Then
                    ' We have an exact match keep looking
                    Continue For
                End If
                If String.Compare(ident.Key, ConvertedIdentifier, ignoreCase:=True) = 0 Then
                    ' If we are here we have seen the variable in a different case so fix it
                    If UsedIdentifiers(ident.Key).IsType Then
                        UsedIdentifiers.Add(ConvertedIdentifier, New SymbolTableEntry(_Name:=ConvertedIdentifier, _IsType:=False))
                    Else
                        Dim NewUniqueName As String = If(ConvertedIdentifier.StartsWith("["), ConvertedIdentifier.Replace("[", "").Replace("]", "_Renamed"), $"{ConvertedIdentifier}_Renamed")
                        UsedIdentifiers.Add(ConvertedIdentifier, New SymbolTableEntry(_Name:=NewUniqueName, _IsType:=QualifiedNameOrTypeName))
                    End If
                    Return SyntaxFactory.Identifier(UsedIdentifiers(ConvertedIdentifier).Name)
                End If
            Next
            UsedIdentifiers.Add(ConvertedIdentifier, New SymbolTableEntry(ConvertedIdentifier, QualifiedNameOrTypeName))
            Return SyntaxFactory.Identifier(ConvertedIdentifier)
        End Function

        ''' <summary>
        ''' CHECK!!!!!!!!
        ''' </summary>
        ''' <param name="node"></param>
        ''' <returns></returns>
        ''' <remarks>Fix handling of AddressOf where is mistakenly added to DirectCast and C# Pointers</remarks>
        Private Shared Function RemodelVariableDeclaration(ByVal declaration As CSS.VariableDeclarationSyntax, ByVal lNodesVisitor As NodesVisitor, ByRef LeadingTrivia As List(Of SyntaxTrivia)) As SeparatedSyntaxList(Of VariableDeclaratorSyntax)
            Dim type As TypeSyntax
            Dim TypeOrAddressOf As VisualBasicSyntaxNode = declaration.Type.Accept(lNodesVisitor).WithConvertedLeadingTriviaFrom(declaration.Type)
            If TypeOrAddressOf.HasLeadingTrivia Then

                LeadingTrivia.AddRange(TypeOrAddressOf.GetLeadingTrivia)
                TypeOrAddressOf = TypeOrAddressOf.WithLeadingTrivia(SyntaxFactory.Whitespace(" "))
            End If
            Dim CollectedCommentTrivia As New List(Of SyntaxTrivia)
            If TypeOrAddressOf.IsKind(SyntaxKind.AddressOfExpression) Then
                type = SyntaxFactory.ParseTypeName(DirectCast(TypeOrAddressOf, UnaryExpressionSyntax).Operand.ToString)
                CollectedCommentTrivia.Add(SyntaxFactory.CommentTrivia(" ' TODO TASK: VB has no direct equivalent to C# Pointer Variables"))
            Else
                type = DirectCast(TypeOrAddressOf, TypeSyntax)
            End If
            Dim declaratorsWithoutInitializers As New List(Of CSS.VariableDeclaratorSyntax)()
            Dim declarators As New List(Of VariableDeclaratorSyntax)
            For i As Integer = 0 To declaration.Variables.Count - 1
                Dim v As CSS.VariableDeclaratorSyntax = declaration.Variables(i)
                If v.Initializer Is Nothing Then
                    declaratorsWithoutInitializers.Add(v.WithTrailingTrivia(CollectedCommentTrivia))
                    Continue For
                Else
                    Dim AsClause As SimpleAsClauseSyntax = If(declaration.Type.IsVar OrElse declaration.Type.IsKind(CSharp.SyntaxKind.RefType), Nothing, SyntaxFactory.SimpleAsClause(type))
                    Dim Value As ExpressionSyntax = DirectCast(v.Initializer.Value.Accept(lNodesVisitor), ExpressionSyntax)
                    If Value.GetLeadingTrivia.ContainsCommentOrDirectiveTrivia Then
                        LeadingTrivia.AddRange(Value.GetLeadingTrivia)
                    End If
                    Dim Initializer As EqualsValueSyntax = SyntaxFactory.EqualsValue(Value.WithLeadingTrivia(SyntaxFactory.ElasticSpace))
                    ' Get the names last to lead with var jsonWriter = new JsonWriter(stringWriter)
                    ' Which should be Dim jsonWriter_Renamed = new JsonWriter(stringWriter)
                    Dim Names As SeparatedSyntaxList(Of ModifiedIdentifierSyntax) = SyntaxFactory.SingletonSeparatedList(ExtractIdentifier(v))
                    Dim Declator As VariableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(
                                                                                                Names,
                                                                                                AsClause,
                                                                                                Initializer
                                                                                                )
                    If Declator.HasTrailingTrivia Then
                        Dim FoundEOL As Boolean = False
                        Dim NonCommentTrailingTrivia As New List(Of SyntaxTrivia)
                        For Each t As SyntaxTrivia In Declator.GetTrailingTrivia
                            Select Case t.Kind
                                Case SyntaxKind.EndOfLineTrivia
                                    FoundEOL = True
                                Case SyntaxKind.CommentTrivia
                                    CollectedCommentTrivia.Add(t)
                                Case SyntaxKind.WhitespaceTrivia
                                    CollectedCommentTrivia.Add(t)
                                    NonCommentTrailingTrivia.Add(t)
                                Case Else
                                    Stop
                            End Select
                        Next
                        If FoundEOL Then
                            CollectedCommentTrivia.Add(VB_EOLTrivia)
                            Declator = Declator.WithTrailingTrivia(CollectedCommentTrivia)
                            CollectedCommentTrivia.Clear()
                        Else
                            Declator = Declator.WithTrailingTrivia(NonCommentTrailingTrivia)
                        End If
                        If i = declaration.Variables.Count - 1 Then
                            If Not Declator.HasTrailingTrivia OrElse Not Declator.GetTrailingTrivia.Last.IsKind(SyntaxKind.EndOfLineTrivia) Then
                                Declator = Declator.WithAppendedTrailingTrivia(VB_EOLTrivia)
                            End If
                        End If
                    End If
                    declarators.Add(Declator)
                End If
            Next
            If declaratorsWithoutInitializers.Count > 0 Then
                declarators.Insert(0, SyntaxFactory.VariableDeclarator(names:=SyntaxFactory.SeparatedList(declaratorsWithoutInitializers.[Select](AddressOf ExtractIdentifier)), asClause:=SyntaxFactory.SimpleAsClause(type), initializer:=Nothing).WithTrailingTrivia(CollectedCommentTrivia))
            End If
            Return SyntaxFactory.SeparatedList(declarators)
        End Function

        Public Shared Function Convert(ByVal input As CS.CSharpSyntaxNode, ByVal lSemanticModel As SemanticModel, ByVal targetDocument As Document) As VisualBasicSyntaxNode
            Dim visualBasicSyntaxNode1 As VisualBasicSyntaxNode
            SyncLock thisLock
                UsedIdentifierStack.Push(UsedIdentifiers)
                UsedIdentifiers.Clear()
                visualBasicSyntaxNode1 = input.Accept(New NodesVisitor(lSemanticModel, targetDocument))
                UsedIdentifiers = DirectCast(UsedIdentifierStack.Pop, Dictionary(Of String, SymbolTableEntry))
            End SyncLock
            Return visualBasicSyntaxNode1
        End Function

        Public Shared Function ConvertText(ByVal InputText As String, ByVal references As MetadataReference()) As ConversionResult
            UsedIdentifierStack.Push(UsedIdentifiers)
            UsedIdentifiers.Clear()
            If InputText Is Nothing Then
                Throw New ArgumentNullException(NameOf(InputText))
            End If
            If references Is Nothing Then
                Throw New ArgumentNullException(NameOf(references))
            End If

            Dim PreprocessorSymbols As New List(Of String) From {
                "NETSTANDARD2_0",
                "NET46",
                "NETCOREAPP2_0"
            }

            Dim CSharpParseOption As CS.CSharpParseOptions = New CS.CSharpParseOptions(
                languageVersion:=CS.LanguageVersion.Latest,
                documentationMode:=DocumentationMode.Parse,
                kind:=SourceCodeKind.Regular,
                preprocessorSymbols:=PreprocessorSymbols)
            Dim tree As SyntaxTree = CS.SyntaxFactory.ParseSyntaxTree(text:=SourceText.From(InputText),
                                                                      options:=CSharpParseOption)
            Dim CSharpOptions As CS.CSharpCompilationOptions = New CS.CSharpCompilationOptions(
                outputKind:=Nothing,
                reportSuppressedDiagnostics:=False,
                moduleName:=Nothing,
                mainTypeName:=Nothing,
                scriptClassName:=Nothing,
                usings:=Nothing,
                optimizationLevel:=OptimizationLevel.Debug,
                checkOverflow:=False,
                allowUnsafe:=True,
                cryptoKeyContainer:=Nothing,
                cryptoKeyFile:=Nothing,
                cryptoPublicKey:=Nothing,
                delaySign:=Nothing,
                platform:=Platform.AnyCpu,
                generalDiagnosticOption:=ReportDiagnostic.Default,
                warningLevel:=4,
                specificDiagnosticOptions:=Nothing,
                concurrentBuild:=True,
                deterministic:=False,
                xmlReferenceResolver:=Nothing,
                sourceReferenceResolver:=Nothing,
                metadataReferenceResolver:=Nothing,
                assemblyIdentityComparer:=Nothing,
                strongNameProvider:=Nothing,
                publicSign:=False)
            Dim compilation As Compilation = CS.CSharpCompilation.Create(assemblyName:=NameOf(Conversion), syntaxTrees:={tree}, references:=references, options:=CSharpOptions)
            Try
                Dim visualBasicSyntaxNode1 As VisualBasicSyntaxNode = Convert(input:=DirectCast(tree.GetRoot(), CS.CSharpSyntaxNode), lSemanticModel:=compilation.GetSemanticModel(tree, ignoreAccessibility:=True), targetDocument:=Nothing)
                UsedIdentifiers = DirectCast(UsedIdentifierStack.Pop, Dictionary(Of String, SymbolTableEntry))
                Return New ConversionResult(visualBasicSyntaxNode1, LanguageNames.CSharp, LanguageNames.VisualBasic)
            Catch ex As Exception
                Return New ConversionResult(ex)
            End Try
        End Function
        Public Shared Function ConvertToken(ByVal t As CS.SyntaxKind, ByVal IsModule As Boolean, ByVal Optional context As TokenContext = TokenContext.[Global]) As SyntaxKind
            Select Case t
                Case CS.SyntaxKind.None
                    Return SyntaxKind.None
                ' built-in types
                Case CS.SyntaxKind.BoolKeyword
                    Return SyntaxKind.BooleanKeyword
                Case CS.SyntaxKind.ByteKeyword
                    Return SyntaxKind.ByteKeyword
                Case CS.SyntaxKind.SByteKeyword
                    Return SyntaxKind.SByteKeyword
                Case CS.SyntaxKind.ShortKeyword
                    Return SyntaxKind.ShortKeyword
                Case CS.SyntaxKind.UShortKeyword
                    Return SyntaxKind.UShortKeyword
                Case CS.SyntaxKind.IntKeyword
                    Return SyntaxKind.IntegerKeyword
                Case CS.SyntaxKind.UIntKeyword
                    Return SyntaxKind.UIntegerKeyword
                Case CS.SyntaxKind.LongKeyword
                    Return SyntaxKind.LongKeyword
                Case CS.SyntaxKind.ULongKeyword
                    Return SyntaxKind.ULongKeyword
                Case CS.SyntaxKind.DoubleKeyword
                    Return SyntaxKind.DoubleKeyword
                Case CS.SyntaxKind.FloatKeyword
                    Return SyntaxKind.SingleKeyword
                Case CS.SyntaxKind.DecimalKeyword
                    Return SyntaxKind.DecimalKeyword
                Case CS.SyntaxKind.StringKeyword
                    Return SyntaxKind.StringKeyword
                Case CS.SyntaxKind.CharKeyword
                    Return SyntaxKind.CharKeyword
                Case CS.SyntaxKind.VoidKeyword
                    ' not supported
                    If context = TokenContext.XMLComment Then
                        Return SyntaxKind.NothingKeyword
                    End If
                    Return SyntaxKind.None
                Case CS.SyntaxKind.ObjectKeyword
                    Return SyntaxKind.ObjectKeyword
                ' literals
                Case CS.SyntaxKind.NullKeyword
                    Return SyntaxKind.NothingKeyword
                Case CS.SyntaxKind.TrueKeyword
                    Return SyntaxKind.TrueKeyword
                Case CS.SyntaxKind.FalseKeyword
                    Return SyntaxKind.FalseKeyword
                Case CS.SyntaxKind.ThisKeyword
                    Return SyntaxKind.MeKeyword
                Case CS.SyntaxKind.BaseKeyword
                    Return SyntaxKind.MyBaseKeyword
                ' modifiers
                Case CS.SyntaxKind.PublicKeyword
                    Return SyntaxKind.PublicKeyword
                Case CS.SyntaxKind.PrivateKeyword
                    Return SyntaxKind.PrivateKeyword
                Case CS.SyntaxKind.InternalKeyword
                    Return SyntaxKind.FriendKeyword
                Case CS.SyntaxKind.ProtectedKeyword
                    Return SyntaxKind.ProtectedKeyword
                Case CS.SyntaxKind.StaticKeyword
                    Return If(IsModule, SyntaxKind.None, SyntaxKind.SharedKeyword)
                Case CS.SyntaxKind.ReadOnlyKeyword
                    Return SyntaxKind.ReadOnlyKeyword
                Case CS.SyntaxKind.SealedKeyword
                    Return If(context = TokenContext.[Global] OrElse context = TokenContext.Class, SyntaxKind.NotInheritableKeyword, SyntaxKind.NotOverridableKeyword)
                Case CS.SyntaxKind.ConstKeyword
                    If context = TokenContext.Readonly Then
                        Return SyntaxKind.ReadOnlyKeyword
                    End If
                    Return SyntaxKind.ConstKeyword
                Case CS.SyntaxKind.OverrideKeyword
                    Return SyntaxKind.OverridesKeyword
                Case CS.SyntaxKind.AbstractKeyword
                    Return If(context = TokenContext.[Global] OrElse context = TokenContext.Class, SyntaxKind.MustInheritKeyword, SyntaxKind.MustOverrideKeyword)
                Case CS.SyntaxKind.VirtualKeyword
                    Return SyntaxKind.OverridableKeyword
                Case CS.SyntaxKind.RefKeyword
                    Return SyntaxKind.ByRefKeyword
                Case CS.SyntaxKind.InKeyword
                    Return SyntaxKind.ByValKeyword
                Case CS.SyntaxKind.OutKeyword
                    Return SyntaxKind.ByRefKeyword
                Case CS.SyntaxKind.PartialKeyword
                    Return SyntaxKind.PartialKeyword
                Case CS.SyntaxKind.AsyncKeyword
                    Return SyntaxKind.AsyncKeyword
                Case CS.SyntaxKind.ExternKeyword
                    ' not supported
                    Return SyntaxKind.None
                Case CS.SyntaxKind.UnsafeKeyword
                    Return SyntaxKind.None
                Case CS.SyntaxKind.NewKeyword
                    Return SyntaxKind.ShadowsKeyword
                Case CS.SyntaxKind.ParamsKeyword
                    Return SyntaxKind.ParamArrayKeyword
                ' others
                Case CS.SyntaxKind.AscendingKeyword
                    Return SyntaxKind.AscendingKeyword
                Case CS.SyntaxKind.DescendingKeyword
                    Return SyntaxKind.DescendingKeyword
                Case CS.SyntaxKind.AwaitKeyword
                    Return SyntaxKind.AwaitKeyword
                ' expressions
                Case CS.SyntaxKind.AddExpression
                    Return SyntaxKind.AddExpression
                Case CS.SyntaxKind.SubtractExpression
                    Return SyntaxKind.SubtractExpression
                Case CS.SyntaxKind.MultiplyExpression
                    Return SyntaxKind.MultiplyExpression
                Case CS.SyntaxKind.DivideExpression
                    Return SyntaxKind.DivideExpression
                Case CS.SyntaxKind.ModuloExpression
                    Return SyntaxKind.ModuloExpression
                Case CS.SyntaxKind.LeftShiftExpression
                    Return SyntaxKind.LeftShiftExpression
                Case CS.SyntaxKind.RightShiftExpression
                    Return SyntaxKind.RightShiftExpression
                Case CS.SyntaxKind.LogicalOrExpression
                    Return SyntaxKind.OrElseExpression
                Case CS.SyntaxKind.LogicalAndExpression
                    Return SyntaxKind.AndAlsoExpression
                Case CS.SyntaxKind.BitwiseOrExpression
                    Return SyntaxKind.OrExpression
                Case CS.SyntaxKind.BitwiseAndExpression
                    Return SyntaxKind.AndExpression
                Case CS.SyntaxKind.ExclusiveOrExpression
                    Return SyntaxKind.ExclusiveOrExpression
                Case CS.SyntaxKind.EqualsExpression
                    Return SyntaxKind.EqualsExpression
                Case CS.SyntaxKind.NotEqualsExpression
                    Return SyntaxKind.NotEqualsExpression
                Case CS.SyntaxKind.LessThanExpression
                    Return SyntaxKind.LessThanExpression
                Case CS.SyntaxKind.LessThanOrEqualExpression
                    Return SyntaxKind.LessThanOrEqualExpression
                Case CS.SyntaxKind.GreaterThanExpression
                    Return SyntaxKind.GreaterThanExpression
                Case CS.SyntaxKind.GreaterThanOrEqualExpression
                    Return SyntaxKind.GreaterThanOrEqualExpression
                Case CS.SyntaxKind.SimpleAssignmentExpression
                    Return SyntaxKind.SimpleAssignmentStatement
                Case CS.SyntaxKind.AddAssignmentExpression
                    Return SyntaxKind.AddAssignmentStatement
                Case CS.SyntaxKind.SubtractAssignmentExpression
                    Return SyntaxKind.SubtractAssignmentStatement
                Case CS.SyntaxKind.MultiplyAssignmentExpression
                    Return SyntaxKind.MultiplyAssignmentStatement
                Case CS.SyntaxKind.DivideAssignmentExpression
                    Return SyntaxKind.DivideAssignmentStatement
                Case CS.SyntaxKind.ModuloAssignmentExpression
                    Return SyntaxKind.ModuloExpression
                Case CS.SyntaxKind.AndAssignmentExpression
                    Return SyntaxKind.AndExpression
                Case CS.SyntaxKind.ExclusiveOrAssignmentExpression
                    Return SyntaxKind.ExclusiveOrExpression
                Case CS.SyntaxKind.OrAssignmentExpression
                    Return SyntaxKind.OrExpression
                Case CS.SyntaxKind.LeftShiftAssignmentExpression
                    Return SyntaxKind.LeftShiftAssignmentStatement
                Case CS.SyntaxKind.RightShiftAssignmentExpression
                    Return SyntaxKind.RightShiftAssignmentStatement
                Case CS.SyntaxKind.UnaryPlusExpression
                    Return SyntaxKind.UnaryPlusExpression
                Case CS.SyntaxKind.UnaryMinusExpression
                    Return SyntaxKind.UnaryMinusExpression
                Case CS.SyntaxKind.BitwiseNotExpression
                    Return SyntaxKind.NotExpression
                Case CS.SyntaxKind.LogicalNotExpression
                    Return SyntaxKind.NotExpression
                Case CS.SyntaxKind.PreIncrementExpression
                    Return SyntaxKind.AddAssignmentStatement
                Case CS.SyntaxKind.PreDecrementExpression
                    Return SyntaxKind.SubtractAssignmentStatement
                Case CS.SyntaxKind.PostIncrementExpression
                    Return SyntaxKind.AddAssignmentStatement
                Case CS.SyntaxKind.PostDecrementExpression
                    Return SyntaxKind.SubtractAssignmentStatement
                Case CS.SyntaxKind.PlusPlusToken
                    Return SyntaxKind.PlusToken
                Case CS.SyntaxKind.MinusMinusToken
                    Return SyntaxKind.MinusToken
                Case CS.SyntaxKind.AddressOfExpression
                    Return SyntaxKind.AddressOfExpression
                Case CS.SyntaxKind.PointerIndirectionExpression
                    Return CType(1999, SyntaxKind)
                Case CS.SyntaxKind.VolatileKeyword
                    ' Function VolatileRead(Of T)(ByRef Address As T) As T
                    '       VolatileRead = Address
                    '       Threading.Thread.MemoryBarrier()
                    ' End Function

                    'Sub VolatileWrite(Of T)(ByRef Address As T, ByVal Value As T)
                    '    Threading.Thread.MemoryBarrier()
                    '    Address = Value
                    'End Sub
                    Return SyntaxKind.None
            End Select

            Throw New NotSupportedException(t & " is not supported!")
        End Function
    End Class
End Namespace
