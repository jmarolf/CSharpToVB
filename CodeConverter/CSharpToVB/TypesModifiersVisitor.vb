Option Explicit On
Option Infer Off
Option Strict On
Imports IVisualBasicCode.CodeConverter.Util
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports ArrayRankSpecifierSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.ArrayRankSpecifierSyntax
Imports CS = Microsoft.CodeAnalysis.CSharp
Imports CSS = Microsoft.CodeAnalysis.CSharp.Syntax
Imports SyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports SyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports TypeParameterConstraintClauseSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeParameterConstraintClauseSyntax
Imports TypeParameterListSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeParameterListSyntax
Imports TypeParameterSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeParameterSyntax
Imports TypeSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax

Namespace IVisualBasicCode.CodeConverter.VB

    Partial Public Class CSharpConverter

        Partial Protected Friend Class NodesVisitor
            Inherits CS.CSharpSyntaxVisitor(Of VisualBasicSyntaxNode)
            Private ReadOnly AsKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.AsKeyword)
            Private ReadOnly CloseParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseParenToken)
            Private ReadOnly CommaToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CommaToken)
            Private ReadOnly OfKeyword As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OfKeyword).WithTrailingTrivia(SyntaxFactory.Space)
            Private ReadOnly OpenParenToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenParenToken)
            Private Shared Function ConvertToTypeString(type As ITypeSymbol) As String
                Dim TypeString As String = type.ToString
                Dim IndexOfLessThan As Integer = TypeString.IndexOf("<")
                If IndexOfLessThan > 0 Then
                    Return SyntaxFactory.ParseTypeName(TypeString.Left(IndexOfLessThan)).ToString & TypeString.Substring(IndexOfLessThan).Replace("<", "(Of ").Replace(">", ")")
                End If
                Return ConvertToType(TypeString).ToString
            End Function

            Private Shared Function FindClauseForParameter(ByVal node As CSS.TypeParameterSyntax) As CSS.TypeParameterConstraintClauseSyntax
                Dim clauses As SyntaxList(Of CSS.TypeParameterConstraintClauseSyntax)
                Dim parentBlock As SyntaxNode = node.Parent.Parent
                If TypeOf parentBlock Is CSS.StructDeclarationSyntax Then
                    Dim s As CSS.StructDeclarationSyntax = DirectCast(parentBlock, CSS.StructDeclarationSyntax)
                    Return CS.SyntaxFactory.TypeParameterConstraintClause(CS.SyntaxFactory.IdentifierName(s.TypeParameterList.Parameters(0).Identifier.Text), Nothing)
                Else
                    clauses = parentBlock.TypeSwitch(
                    Function(ByVal m As CSS.MethodDeclarationSyntax) m.ConstraintClauses,
                    Function(ByVal c As CSS.ClassDeclarationSyntax) c.ConstraintClauses,
                    Function(ByVal d As CSS.DelegateDeclarationSyntax) d.ConstraintClauses,
                    Function(ByVal i As CSS.InterfaceDeclarationSyntax) i.ConstraintClauses,
                    Function(Underscore As SyntaxNode) As SyntaxList(Of CSS.TypeParameterConstraintClauseSyntax)
                        Throw New NotImplementedException($"{Underscore.[GetType]().FullName} not implemented!")
                    End Function)
                    Return clauses.FirstOrDefault(Function(c As CSS.TypeParameterConstraintClauseSyntax) c.Name.ToString() = node.ToString())
                End If
            End Function
            Public Shared Function ConvertToType(TypeString As String) As Syntax.TypeSyntax
                Select Case TypeString.ToLower
                    Case "byte"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ByteKeyword))
                    Case "sbyte"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.SByteKeyword))
                    Case "int"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.IntegerKeyword))
                    Case "uint"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.UIntegerKeyword))
                    Case "short"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ShortKeyword))
                    Case "ushort"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.UShortKeyword))
                    Case "long"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.LongKeyword))
                    Case "ulong"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ULongKeyword))
                    Case "float"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.SingleKeyword))
                    Case "double"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.DoubleKeyword))
                    Case "char"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.CharKeyword))
                    Case "bool"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.BooleanKeyword))
                    Case "object", "var"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ObjectKeyword))
                    Case "string"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.StringKeyword))
                    Case "decimal"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.DecimalKeyword))
                    Case "datetime"
                        Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.DateKeyword))
                    Case Else
                        Return SyntaxFactory.ParseTypeName(TypeString)
                End Select
            End Function

            Public Overrides Function VisitArrayRankSpecifier(ByVal node As CSS.ArrayRankSpecifierSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.ArrayRankSpecifier(OpenParenToken, SyntaxFactory.TokenList(Enumerable.Repeat(CommaToken, node.Rank - 1)), CloseParenToken)
            End Function

            Public Overrides Function VisitArrayType(ByVal node As CSS.ArrayTypeSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.ArrayType(DirectCast(node.ElementType.Accept(Me), TypeSyntax), SyntaxFactory.List(node.RankSpecifiers.[Select](Function(rs As CSS.ArrayRankSpecifierSyntax) DirectCast(rs.Accept(Me), ArrayRankSpecifierSyntax)))).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitClassOrStructConstraint(ByVal node As CSS.ClassOrStructConstraintSyntax) As VisualBasicSyntaxNode
                If node.IsKind(CS.SyntaxKind.ClassConstraint) Then
                    Return SyntaxFactory.ClassConstraint(SyntaxFactory.Token(SyntaxKind.ClassKeyword)).WithConvertedTriviaFrom(node)
                End If
                If node.IsKind(CS.SyntaxKind.StructConstraint) Then
                    Return SyntaxFactory.StructureConstraint(SyntaxFactory.Token(SyntaxKind.StructureKeyword)).WithConvertedTriviaFrom(node)
                End If
                Throw New NotSupportedException()
            End Function

            Public Overrides Function VisitConstructorConstraint(ByVal node As CSS.ConstructorConstraintSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.NewConstraint(SyntaxFactory.Token(SyntaxKind.NewKeyword)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitNullableType(ByVal node As CSS.NullableTypeSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.NullableType(DirectCast(node.ElementType.Accept(Me), TypeSyntax)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitPointerType(node As PointerTypeSyntax) As VisualBasicSyntaxNode
                If node.ToString = "void*" Then
                    Return SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ObjectKeyword)).WithConvertedTriviaFrom(node.Parent)
                End If
                Dim NodeParent As SyntaxNode = node.Parent
                If TypeOf NodeParent Is CSS.CastExpressionSyntax Then
                    Dim OperandWithAmpersand As CSS.ExpressionSyntax = DirectCast(node.Parent, CSS.CastExpressionSyntax).Expression
                    Return SyntaxFactory.AddressOfExpression(SyntaxFactory.ParseExpression(OperandWithAmpersand.ToString).WithConvertedTriviaFrom(OperandWithAmpersand))
                End If
                If TypeOf NodeParent Is CSS.VariableDeclarationSyntax Then
                    Dim Operand As TypeSyntax = DirectCast(node.ElementType.Accept(Me), TypeSyntax)
                    Return SyntaxFactory.AddressOfExpression(SyntaxFactory.ParseExpression(Operand.ToString).WithConvertedTriviaFrom(node.ElementType))
                End If
                If TypeOf NodeParent Is CSS.ParameterSyntax Then
                    Dim Operand As TypeSyntax = DirectCast(node.ElementType.Accept(Me), TypeSyntax)
                    Return Operand.WithConvertedTriviaFrom(node.ElementType)
                End If
                If node.ToString = "int*" Then
                    Return SyntaxFactory.ParseTypeName("Intptr")
                End If
                If node.ToString = "char*" Then
                    Return SyntaxFactory.ParseTypeName("Intptr")
                End If
                Return SyntaxFactory.ParseTypeName("Intptr")
            End Function

            Public Overrides Function VisitPredefinedType(ByVal node As CSS.PredefinedTypeSyntax) As VisualBasicSyntaxNode
                Dim PredefinedType As Syntax.PredefinedTypeSyntax = Nothing
                Try
                    If node.Keyword.ToString = "void" Then
                        Return SyntaxFactory.IdentifierName("void")
                    End If
                    PredefinedType = SyntaxFactory.PredefinedType(SyntaxFactory.Token(ConvertToken(CS.CSharpExtensions.Kind(node.Keyword), IsModule)))
                Catch ex As Exception
                    Stop
                End Try
                Return PredefinedType
            End Function

            Public Overrides Function VisitSimpleBaseType(node As SimpleBaseTypeSyntax) As VisualBasicSyntaxNode
                Dim TypeString As String = node.NormalizeWhitespace.ToString

                Return ConvertToType(TypeString).WithConvertedTriviaFrom(node)
            End Function
            Public Overrides Function VisitTypeConstraint(ByVal node As CSS.TypeConstraintSyntax) As VisualBasicSyntaxNode
                Return SyntaxFactory.TypeConstraint(DirectCast(node.Type.Accept(Me), TypeSyntax)).WithConvertedTriviaFrom(node)
            End Function

            Public Overrides Function VisitTypeParameter(ByVal node As CSS.TypeParameterSyntax) As VisualBasicSyntaxNode
                Dim variance As SyntaxToken = Nothing
                If Not SyntaxTokenExtensions.IsKind(node.VarianceKeyword, CS.SyntaxKind.None) Then
                    variance = SyntaxFactory.Token(If(SyntaxTokenExtensions.IsKind(node.VarianceKeyword, CS.SyntaxKind.InKeyword), SyntaxKind.InKeyword, SyntaxKind.OutKeyword))
                End If

                ' copy generic constraints
                Dim clause As CSS.TypeParameterConstraintClauseSyntax = FindClauseForParameter(node)
                Dim TypeParameterConstraintClause As TypeParameterConstraintClauseSyntax = DirectCast(clause?.Accept(Me), TypeParameterConstraintClauseSyntax)
                If TypeParameterConstraintClause IsNot Nothing AndAlso TypeParameterConstraintClause.Kind = SyntaxKind.TypeParameterMultipleConstraintClause Then
                    Dim TypeParameterMultipleConstraintClause As TypeParameterMultipleConstraintClauseSyntax = DirectCast(TypeParameterConstraintClause, TypeParameterMultipleConstraintClauseSyntax)
                    If TypeParameterMultipleConstraintClause.Constraints.Count = 0 Then
                        TypeParameterConstraintClause = Nothing
                    End If
                End If
                Dim TypeParameterSyntax As TypeParameterSyntax = SyntaxFactory.TypeParameter(variance, GenerateSafeVBToken(node.Identifier, IsQualifiedName:=False), TypeParameterConstraintClause).WithConvertedTriviaFrom(node)
                Return TypeParameterSyntax
            End Function

            Public Overrides Function VisitTypeParameterConstraintClause(ByVal node As CSS.TypeParameterConstraintClauseSyntax) As VisualBasicSyntaxNode
                Dim Braces As (OpenBrace As SyntaxToken, CloseBrace As SyntaxToken) = node.GetBraces
                Dim OpenBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.OpenBraceToken).WithConvertedTriviaFrom(Braces.OpenBrace)
                Dim CloseBraceToken As SyntaxToken = SyntaxFactory.Token(SyntaxKind.CloseBraceToken).WithConvertedTriviaFrom(Braces.CloseBrace)
                If node.Constraints.Count = 1 Then
                    Return SyntaxFactory.TypeParameterSingleConstraintClause(AsKeyword, DirectCast(node.Constraints(0).Accept(Me), ConstraintSyntax))
                End If
                Dim Constraints As SeparatedSyntaxList(Of ConstraintSyntax) = SyntaxFactory.SeparatedList(node.Constraints.[Select](Function(c As CSS.TypeParameterConstraintSyntax) DirectCast(c.Accept(Me), ConstraintSyntax)))
                Return SyntaxFactory.TypeParameterMultipleConstraintClause(AsKeyword, OpenBraceToken, Constraints, CloseBraceToken)
            End Function

            Public Overrides Function VisitTypeParameterList(ByVal node As CSS.TypeParameterListSyntax) As VisualBasicSyntaxNode
                Dim Nodes As New List(Of TypeParameterSyntax)
                Dim Separators As New List(Of SyntaxToken)
                Dim CS_Separators As New List(Of SyntaxToken)
                CS_Separators.AddRange(node.Parameters.GetSeparators)
                Dim FinalTrailingTrivia As New List(Of SyntaxTrivia)
                For i As Integer = 0 To node.Parameters.Count - 2
                    Dim p As CSS.TypeParameterSyntax = node.Parameters(i)
                    Dim ItemWithTrivia As TypeParameterSyntax = DirectCast(p.Accept(Me), TypeParameterSyntax)
                    FinalTrailingTrivia.AddRange(ExtractComments(ItemWithTrivia.GetLeadingTrivia, Leading:=True))
                    FinalTrailingTrivia.AddRange(ExtractComments(ItemWithTrivia.GetTrailingTrivia, Leading:=False))
                    Nodes.Add(ItemWithTrivia.WithLeadingTrivia(SyntaxFactory.ElasticSpace).WithTrailingTrivia(SyntaxFactory.ElasticSpace))
                    Separators.Add(CommaToken.WithConvertedTriviaFrom(CS_Separators(i)))
                Next
                Nodes.Add(DirectCast(node.Parameters.Last.Accept(Me).WithConvertedTrailingTriviaFrom(node.Parameters.Last), TypeParameterSyntax))
                Dim SeparatedList As SeparatedSyntaxList(Of TypeParameterSyntax) = SyntaxFactory.SeparatedList(Nodes, Separators)
                Dim TypeParameterListSyntax As TypeParameterListSyntax = SyntaxFactory.TypeParameterList(openParenToken:=OpenParenToken,
                                                                                                         ofKeyword:=OfKeyword.WithConvertedTriviaFrom(node.LessThanToken),
                                                                                                         parameters:=SeparatedList,
                                                                                                         closeParenToken:=CloseParenToken.WithConvertedTriviaFrom(node.GreaterThanToken).WithAppendedTrailingTrivia(FinalTrailingTrivia))
                Return TypeParameterListSyntax
            End Function

        End Class
    End Class
End Namespace