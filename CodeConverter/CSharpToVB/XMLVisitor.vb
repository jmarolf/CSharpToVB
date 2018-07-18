Imports IVisualBasicCode.CodeConverter.Util
Imports IVisualBasicCode.CodeConverter.VB.CSharpConverter
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax

Namespace IVisualBasicCode.CodeConverter.VB

    Friend Class XMLVisitor
        Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)

        Private ReadOnly mNodesVisitor As NodesVisitor
        Private ReadOnly semanticModel As SemanticModel
        Public Sub New(ByVal semanticModel As SemanticModel, ByVal nodesVisitor As NodesVisitor)
            mNodesVisitor = nodesVisitor
            Me.semanticModel = semanticModel
        End Sub

        Private Shared Function GetVBOperatorToken(ByVal op As String) As SyntaxToken
            Select Case op
                Case "=="
                    Return SyntaxFactory.Token(SyntaxKind.EqualsToken)
                Case "!="
                    Return SyntaxFactory.Token(SyntaxKind.LessThanGreaterThanToken)
                Case ">"
                    Return SyntaxFactory.Token(SyntaxKind.GreaterThanToken)
                Case ">="
                    Return SyntaxFactory.Token(SyntaxKind.GreaterThanEqualsToken)
                Case "<"
                    Return SyntaxFactory.Token(SyntaxKind.LessThanToken)
                Case "<="
                    Return SyntaxFactory.Token(SyntaxKind.LessThanEqualsToken)
                Case "|"
                    Return SyntaxFactory.Token(SyntaxKind.OrKeyword)
                Case "||"
                    Return SyntaxFactory.Token(SyntaxKind.OrElseKeyword)
                Case "&"
                    Return SyntaxFactory.Token(SyntaxKind.AndKeyword)
                Case "&&"
                    Return SyntaxFactory.Token(SyntaxKind.AndAlsoKeyword)
                Case "+"
                    Return SyntaxFactory.Token(SyntaxKind.PlusToken)
                Case "-"
                    Return SyntaxFactory.Token(SyntaxKind.MinusToken)
                Case "*"
                    Return SyntaxFactory.Token(SyntaxKind.AsteriskToken)
                Case "/"
                    Return SyntaxFactory.Token(SyntaxKind.SlashToken)
                Case "%"
                    Return SyntaxFactory.Token(SyntaxKind.ModKeyword)
                Case "="
                    Return SyntaxFactory.Token(SyntaxKind.EqualsToken)
                Case "+="
                    Return SyntaxFactory.Token(SyntaxKind.PlusEqualsToken)
                Case "-="
                    Return SyntaxFactory.Token(SyntaxKind.MinusEqualsToken)
                Case "!"
                    Return SyntaxFactory.Token(SyntaxKind.NotKeyword)
                Case "~"
                    Return SyntaxFactory.Token(SyntaxKind.NotKeyword)
                Case Else

            End Select

            Throw New ArgumentOutOfRangeException(NameOf(op))
        End Function

        Private Function GatherAttributes(Attributes As SyntaxList(Of XmlAttributeSyntax)) As SyntaxList(Of Syntax.XmlNodeSyntax)
            Dim VBAttributes As New SyntaxList(Of Syntax.XmlNodeSyntax)
            For Each a As XmlAttributeSyntax In Attributes
                VBAttributes = VBAttributes.Add(DirectCast(a.Accept(Me), Syntax.XmlNodeSyntax))
            Next
            Return VBAttributes
        End Function
        Public Overrides Function DefaultVisit(node As SyntaxNode) As VisualBasicSyntaxNode
            Return MyBase.DefaultVisit(node)
        End Function
        Public Overrides Function VisitConversionOperatorMemberCref(node As ConversionOperatorMemberCrefSyntax) As VisualBasicSyntaxNode
            Return MyBase.VisitConversionOperatorMemberCref(node)
        End Function

        Public Overrides Function VisitCrefBracketedParameterList(node As CrefBracketedParameterListSyntax) As VisualBasicSyntaxNode
            Return MyBase.VisitCrefBracketedParameterList(node)
        End Function

        Public Overrides Function VisitCrefParameter(node As CrefParameterSyntax) As VisualBasicSyntaxNode
            Return node.Type.Accept(Me)
        End Function

        Public Overrides Function VisitCrefParameterList(node As CrefParameterListSyntax) As VisualBasicSyntaxNode
            Return MyBase.VisitCrefParameterList(node)
        End Function

        Public Overrides Function VisitGenericName(node As GenericNameSyntax) As VisualBasicSyntaxNode
            Dim Identifier As SyntaxToken = SyntaxFactory.Identifier(node.Identifier.ToString)
            Dim TypeList As New List(Of Syntax.TypeSyntax)
            For Each a As TypeSyntax In node.TypeArgumentList.Arguments
                Dim TypeIdentifier As Syntax.TypeSyntax = DirectCast(a.Accept(Me), Syntax.TypeSyntax)
                TypeList.Add(TypeIdentifier)
            Next
            Return SyntaxFactory.GenericName(Identifier, SyntaxFactory.TypeArgumentList(TypeList.ToArray))
        End Function
        Public Overrides Function VisitIdentifierName(node As IdentifierNameSyntax) As VisualBasicSyntaxNode
            Dim Identifier As Syntax.IdentifierNameSyntax = SyntaxFactory.IdentifierName(node.Identifier.ToString)
            Return Identifier
        End Function

        ''' <summary>
        ''' from <see cref="SourceText.From(Stream, Encoding, SourceHashAlgorithm, bool)"/> in two ways
        ''' </summary>
        ''' <param name="node"></param>
        ''' <returns></returns>
        Public Overrides Function VisitNameMemberCref(node As NameMemberCrefSyntax) As VisualBasicSyntaxNode
            Dim Name As VisualBasicSyntaxNode = node.Name.Accept(Me)
            Dim CrefParameters As New List(Of Syntax.CrefSignaturePartSyntax)
            Dim Signature As Syntax.CrefSignatureSyntax = Nothing
            If node.Parameters IsNot Nothing Then
                For Each p As CrefParameterSyntax In node.Parameters.Parameters
                    Dim TypeSyntax1 As Syntax.TypeSyntax = DirectCast(p.Accept(Me), Syntax.TypeSyntax)
                    CrefParameters.Add(SyntaxFactory.CrefSignaturePart(modifier:=Nothing, TypeSyntax1))
                Next
                Signature = SyntaxFactory.CrefSignature(CrefParameters.ToArray)
            End If
            Return SyntaxFactory.CrefReference(DirectCast(Name, Syntax.TypeSyntax), signature:=Signature, asClause:=Nothing)
        End Function

        Public Overrides Function VisitOperatorMemberCref(node As OperatorMemberCrefSyntax) As VisualBasicSyntaxNode
            Dim CrefOperator As SyntaxToken = GetVBOperatorToken(node.OperatorToken.ValueText)
            Return SyntaxFactory.CrefOperatorReference(CrefOperator.WithLeadingTrivia(SyntaxFactory.Space))
        End Function

        Public Overrides Function VisitPredefinedType(node As PredefinedTypeSyntax) As VisualBasicSyntaxNode
            Dim Kind As SyntaxKind = ConvertToken(CS.CSharpExtensions.Kind(node.Keyword), IsModule:=True, context:=TokenContext.XMLComment)
            Select Case Kind
                Case SyntaxKind.None
                    Return SyntaxFactory.ParseTypeName(node.ToString)
                Case SyntaxKind.NothingKeyword
                    Return SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(Kind))
                Case Else
                    Dim PredefinedType As Syntax.TypeSyntax = SyntaxFactory.PredefinedType(SyntaxFactory.Token(Kind))
                    Return PredefinedType
            End Select
        End Function

        Public Overrides Function VisitQualifiedCref(QualifiedCref As QualifiedCrefSyntax) As VisualBasicSyntaxNode
            Dim IdentifierOrTypeName As VisualBasicSyntaxNode = QualifiedCref.Container.Accept(Me)
            Dim DotToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.DotToken)
            Dim Value As Syntax.CrefReferenceSyntax = DirectCast(QualifiedCref.Member.Accept(Me), Syntax.CrefReferenceSyntax)
            Dim Identifier As Syntax.NameSyntax
            Identifier = If(TypeOf IdentifierOrTypeName Is Syntax.NameSyntax, DirectCast(IdentifierOrTypeName, Syntax.NameSyntax), SyntaxFactory.IdentifierName(IdentifierOrTypeName.ToString))
            Dim QualifiedNameSyntax As Syntax.QualifiedNameSyntax = SyntaxFactory.QualifiedName(left:=Identifier, dotToken:=DotToken, right:=DirectCast(Value.Name, Syntax.SimpleNameSyntax))
            If Value.Signature Is Nothing Then
                Return QualifiedNameSyntax
            End If
            Return SyntaxFactory.CrefReference(QualifiedNameSyntax, Value.Signature, Nothing)
        End Function
        Public Overrides Function VisitQualifiedName(node As QualifiedNameSyntax) As VisualBasicSyntaxNode
            Return SyntaxFactory.QualifiedName(DirectCast(node.Left.Accept(Me), Syntax.NameSyntax), DirectCast(node.Right.Accept(Me), Syntax.SimpleNameSyntax))
        End Function
        Public Overrides Function VisitTypeCref(node As TypeCrefSyntax) As VisualBasicSyntaxNode
            Return node.Type.Accept(Me)
        End Function

        Public Overrides Function VisitXmlCDataSection(node As CSS.XmlCDataSectionSyntax) As VisualBasicSyntaxNode
            Dim BeginCDataToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.BeginCDataToken)
            Dim TextTokens As SyntaxTokenList = TranslateTokenList(node.TextTokens)
            Dim EndCDataToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.EndCDataToken)
            Return SyntaxFactory.XmlCDataSection(BeginCDataToken, TextTokens, EndCDataToken)
        End Function

        Public Overrides Function VisitXmlComment(node As CSS.XmlCommentSyntax) As VisualBasicSyntaxNode
            Return MyBase.VisitXmlComment(node).WithConvertedTriviaFrom(node)
        End Function
        Public Overrides Function VisitXmlCrefAttribute(node As CSS.XmlCrefAttributeSyntax) As VisualBasicSyntaxNode
            Dim Name As Syntax.XmlNameSyntax = DirectCast(node.Name.Accept(Me), Syntax.XmlNameSyntax)
            Dim StartQuoteToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken)
            Dim EndQuoteToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken)

            Dim cref As VisualBasicSyntaxNode = node.Cref.Accept(Me)
            Dim SyntaxTokens As New SyntaxTokenList
            SyntaxTokens = SyntaxTokens.AddRange(cref.DescendantTokens)
            Dim Value As Syntax.XmlNodeSyntax = SyntaxFactory.XmlString(
                            SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken),
                            SyntaxTokens,
                            SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken)).WithConvertedTriviaFrom(node)
            Return SyntaxFactory.XmlAttribute(Name, Value)
        End Function
        Public Overrides Function VisitXmlElement(node As CSS.XmlElementSyntax) As VisualBasicSyntaxNode
            Dim StartTag As Syntax.XmlElementStartTagSyntax = DirectCast(node.StartTag.Accept(Me).WithConvertedTriviaFrom(node.StartTag), Syntax.XmlElementStartTagSyntax)
            Dim Content As New SyntaxList(Of Syntax.XmlNodeSyntax)
            Dim EndTag As Syntax.XmlElementEndTagSyntax
            Dim NoEndTag As Boolean = False
            If node.EndTag.Name.LocalName.ValueText.IsEmptyNullOrWhitespace Then
                EndTag = SyntaxFactory.XmlElementEndTag(DirectCast(StartTag.Name, Syntax.XmlNameSyntax).WithConvertedTriviaFrom(node.EndTag))
                NoEndTag = True
            Else
                EndTag = SyntaxFactory.XmlElementEndTag(DirectCast(node.EndTag.Name.Accept(Me), Syntax.XmlNameSyntax))
            End If
            For i As Integer = 0 To node.Content.Count - 1
                Dim C As XmlNodeSyntax = node.Content(i)
                Dim Node1 As Syntax.XmlNodeSyntax = DirectCast(C.Accept(Me).WithConvertedTriviaFrom(C), Syntax.XmlNodeSyntax)
                If NoEndTag Then
                    Dim LastToken As SyntaxToken = Node1.GetLastToken
                    If LastToken.ValueText.IsNewLine Then
                        Node1 = Node1.ReplaceToken(LastToken, CType(Nothing, SyntaxToken))
                    End If
                End If
                Content = Content.Add(Node1)
            Next
            Dim EndTagLeadingTrivia As New List(Of SyntaxTrivia)
            EndTagLeadingTrivia.AddRange(ConvertTrivia(node.EndTag.GetLeadingTrivia))
            If EndTagLeadingTrivia.Count = 1 AndAlso node.EndTag.HasLeadingTrivia AndAlso Content.Last.GetLastToken.IsKind(SyntaxKind.DocumentationCommentLineBreakToken) Then
                EndTag = EndTag.WithPrependedLeadingTrivia(StartTag.GetLeadingTrivia)
            End If
            If EndTag.GetLeadingTrivia.Count > 1 AndAlso Content.Count > 0 AndAlso Content(0).GetFirstToken.GetNextToken.LeadingTrivia.Count > 0 Then
                EndTag = EndTag.WithLeadingTrivia(Content(0).GetFirstToken.GetNextToken.LeadingTrivia)
            End If
            Dim XmlElement As Syntax.XmlElementSyntax = SyntaxFactory.XmlElement(StartTag, Content, EndTag)
            Return XmlElement
        End Function
        ''' <summary>
        ''' Represents a non-terminal node in the syntax tree. This is the language agnostic equivalent of <see
        ''' cref="T:Microsoft.CodeAnalysis.CSharp.SyntaxNode"/>  and <see cref="T:Microsoft.CodeAnalysis.VisualBasic.SyntaxNode"/>.
        ''' </summary>
        Public Overrides Function VisitXmlElementEndTag(node As CSS.XmlElementEndTagSyntax) As VisualBasicSyntaxNode
            Return SyntaxFactory.XmlElementEndTag(DirectCast(node.Name.Accept(Me), Syntax.XmlNameSyntax))
        End Function

        Public Overrides Function VisitXmlElementStartTag(node As CSS.XmlElementStartTagSyntax) As VisualBasicSyntaxNode
            Dim Attributes As SyntaxList(Of Syntax.XmlNodeSyntax) = GatherAttributes(node.Attributes)
            Return SyntaxFactory.XmlElementStartTag(DirectCast(node.Name.Accept(Me), Syntax.XmlNodeSyntax), Attributes).WithConvertedTriviaFrom(node)
        End Function

        Public Overrides Function VisitXmlEmptyElement(node As CSS.XmlEmptyElementSyntax) As VisualBasicSyntaxNode
            Dim Name As Syntax.XmlNodeSyntax = DirectCast(node.Name.Accept(Me), Syntax.XmlNodeSyntax)
            Dim Attributes As SyntaxList(Of Syntax.XmlNodeSyntax) = GatherAttributes(node.Attributes)
            Return SyntaxFactory.XmlEmptyElement(Name, Attributes).WithConvertedTriviaFrom(node)
        End Function
        Public Overrides Function VisitXmlName(node As CSS.XmlNameSyntax) As VisualBasicSyntaxNode
            Dim Prefix As Syntax.XmlPrefixSyntax
            Prefix = If(node.Prefix Is Nothing, Nothing, DirectCast(node.Prefix.Accept(Me), Syntax.XmlPrefixSyntax))
            Dim localName As SyntaxToken = SyntaxFactory.XmlNameToken(node.LocalName.ValueText, Nothing)
            Return SyntaxFactory.XmlName(Prefix, localName)
        End Function

        Public Overrides Function VisitXmlNameAttribute(node As CSS.XmlNameAttributeSyntax) As VisualBasicSyntaxNode
            Dim Name As Syntax.XmlNodeSyntax = DirectCast(node.Name.Accept(Me), Syntax.XmlNodeSyntax).WithConvertedLeadingTriviaFrom(node.Name)
            Dim ValueString As String = node.Identifier.ToString
            Dim Value As Syntax.XmlNodeSyntax = SyntaxFactory.XmlString(
                                    SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken),
                                    SyntaxTokenList.Create(
                                    SyntaxFactory.XmlTextLiteralToken(ValueString, ValueString)),
                                    SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken))
            Return SyntaxFactory.XmlAttribute(Name, Value).WithConvertedTriviaFrom(node).WithConvertedTriviaFrom(node)
        End Function

        Public Overrides Function VisitXmlPrefix(node As CSS.XmlPrefixSyntax) As VisualBasicSyntaxNode
            Return MyBase.VisitXmlPrefix(node).WithConvertedTriviaFrom(node)
        End Function

        Public Overrides Function VisitXmlProcessingInstruction(node As CSS.XmlProcessingInstructionSyntax) As VisualBasicSyntaxNode
            Return MyBase.VisitXmlProcessingInstruction(node).WithConvertedTriviaFrom(node)
        End Function

        Public Overrides Function VisitXmlText(node As CSS.XmlTextSyntax) As VisualBasicSyntaxNode
            Dim TextTokens As SyntaxTokenList = TranslateTokenList(node.TextTokens)
            Dim XmlText As Syntax.XmlTextSyntax = SyntaxFactory.XmlText(TextTokens)
            Return XmlText
        End Function

        Public Overrides Function VisitXmlTextAttribute(node As XmlTextAttributeSyntax) As VisualBasicSyntaxNode
            Dim Name As Syntax.XmlNodeSyntax = DirectCast(node.Name.Accept(Me).WithConvertedTriviaFrom(node.Name), Syntax.XmlNodeSyntax)
            Dim TextTokens As SyntaxTokenList = TranslateTokenList(node.TextTokens)
            Dim XmlText As Syntax.XmlTextSyntax = SyntaxFactory.XmlText(TextTokens)
            Dim Value As Syntax.XmlNodeSyntax = SyntaxFactory.XmlString(
                                    SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken),
                                    SyntaxTokenList.Create(
                                    SyntaxFactory.XmlTextLiteralToken(XmlText.ToString, XmlText.ToString)),
                                    SyntaxFactory.Token(SyntaxKind.DoubleQuoteToken))
            Return SyntaxFactory.XmlAttribute(Name, Value)
        End Function
    End Class
End Namespace