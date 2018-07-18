Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.PooledObjects

Namespace IVisualBasicCode.CodeConverter.Util
    Friend Module SyntaxTokenExtensions
        Private ReadOnly s_childEnumeratorStackPool As New ObjectPool(Of Stack(Of ChildSyntaxList.Enumerator))(Function() New Stack(Of ChildSyntaxList.Enumerator)(), 10)

        <Extension>
        Public Function [With](ByVal token As SyntaxToken, ByVal leading As List(Of SyntaxTrivia), ByVal trailing As List(Of SyntaxTrivia)) As SyntaxToken
            Return token.WithLeadingTrivia(leading).WithTrailingTrivia(trailing)
        End Function

        <Extension>
        Public Function [With](ByVal token As SyntaxToken, ByVal leading As SyntaxTriviaList, ByVal trailing As SyntaxTriviaList) As SyntaxToken
            Return token.WithLeadingTrivia(leading).WithTrailingTrivia(trailing)
        End Function

        ''' <summary>
        ''' Returns the token after this token in the syntax tree.
        ''' </summary>
        ''' <param name="predicate">Delegate applied to each token.  The token is returned if the predicate returns
        ''' true.</param>
        ''' <param name="stepInto">Delegate applied to trivia.  If this delegate is present then trailing trivia is
        ''' included in the search.</param>
        <Extension>
        Public Function GetNextToken(Node As SyntaxToken, predicate As Func(Of SyntaxToken, Boolean), Optional stepInto As Func(Of SyntaxTrivia, Boolean) = Nothing) As SyntaxToken
            If Node = Nothing Then
                Return Nothing
            End If

            Return SyntaxNavigator.Instance.GetNextToken(Node, predicate, stepInto)
        End Function

        <Extension>
        Public Function IsKind(ByVal token As SyntaxToken, ByVal kind1 As CSharp.SyntaxKind, ByVal kind2 As CSharp.SyntaxKind) As Boolean
            Return CSharp.CSharpExtensions.Kind(token) = kind1 OrElse CSharp.CSharpExtensions.Kind(token) = kind2
        End Function

        <Extension>
        Public Function IsKind(ByVal token As SyntaxToken, ByVal kind1 As VisualBasic.SyntaxKind, ByVal kind2 As VisualBasic.SyntaxKind) As Boolean
            Return Microsoft.CodeAnalysis.VisualBasic.VisualBasicExtensions.Kind(token) = kind1 OrElse Microsoft.CodeAnalysis.VisualBasic.VisualBasicExtensions.Kind(token) = kind2
        End Function

        <Extension>
        Public Function IsKind(ByVal token As SyntaxToken, ParamArray ByVal kinds() As Microsoft.CodeAnalysis.CSharp.SyntaxKind) As Boolean
            Return kinds.Contains(Microsoft.CodeAnalysis.CSharp.CSharpExtensions.Kind(token))
        End Function

        <Extension>
        Public Function IsKind(ByVal token As SyntaxToken, ParamArray ByVal kinds() As VisualBasic.SyntaxKind) As Boolean
            Return kinds.Contains(Microsoft.CodeAnalysis.VisualBasic.VisualBasicExtensions.Kind(token))
        End Function

        <Extension>
        Public Function Width(ByVal token As SyntaxToken) As Integer
            Return token.Span.Length
        End Function

        <Extension>
        Public Function WithAppendedTrailingTrivia(ByVal token As SyntaxToken, ByVal trivia As IEnumerable(Of SyntaxTrivia)) As SyntaxToken
            Return token.WithTrailingTrivia(token.TrailingTrivia.Concat(trivia))
        End Function

        <Extension>
        Public Function WithPrependedLeadingTrivia(ByVal token As SyntaxToken, ParamArray ByVal trivia() As SyntaxTrivia) As SyntaxToken
            If trivia.Length = 0 Then
                Return token
            End If

            Return token.WithPrependedLeadingTrivia(DirectCast(trivia, IEnumerable(Of SyntaxTrivia)))
        End Function

        <Extension>
        Public Function WithPrependedLeadingTrivia(ByVal token As SyntaxToken, ByVal trivia As SyntaxTriviaList) As SyntaxToken
            If trivia.Count = 0 Then
                Return token
            End If

            Return token.WithLeadingTrivia(trivia.Concat(token.LeadingTrivia))
        End Function

        <Extension>
        Public Function WithPrependedLeadingTrivia(ByVal token As SyntaxToken, ByVal trivia As IEnumerable(Of SyntaxTrivia)) As SyntaxToken
            Return token.WithPrependedLeadingTrivia(trivia.ToSyntaxTriviaList())
        End Function

    End Module
End Namespace