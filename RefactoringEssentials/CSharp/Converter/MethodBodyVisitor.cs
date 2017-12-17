﻿using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using VBasic = Microsoft.CodeAnalysis.VisualBasic;
using VBSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace RefactoringEssentials.CSharp.Converter
{
	public partial class VisualBasicConverter
	{
		class MethodBodyVisitor : VBasic.VisualBasicSyntaxVisitor<SyntaxList<StatementSyntax>>
		{
			SemanticModel semanticModel;
			NodesVisitor nodesVisitor;
			private readonly Stack<string> withBlockTempVariableNames;

            public bool IsIterator { get; set; }

			public MethodBodyVisitor(SemanticModel semanticModel, NodesVisitor nodesVisitor, Stack<string> withBlockTempVariableNames)
			{
				this.semanticModel = semanticModel;
				this.nodesVisitor = nodesVisitor;
				this.withBlockTempVariableNames = withBlockTempVariableNames;
			}

			public override SyntaxList<StatementSyntax> DefaultVisit(SyntaxNode node)
			{
				throw new NotImplementedException(node.GetType() + " not implemented!");
			}

			public override SyntaxList<StatementSyntax> VisitStopOrEndStatement(VBSyntax.StopOrEndStatementSyntax node)
			{
				var cSharpEquivalent = node.StopOrEndKeyword.IsKind(VBasic.SyntaxKind.StopKeyword) ? "System.Diagnostics.Debugger.Break();"
					: node.StopOrEndKeyword.IsKind(VBasic.SyntaxKind.EndKeyword) ? "System.Environment.Exit(0);"
						: throw new NotImplementedException(node.StopOrEndKeyword.Kind() + " not implemented!");
				return SingleStatement(SyntaxFactory.ParseStatement(cSharpEquivalent)).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitLocalDeclarationStatement(VBSyntax.LocalDeclarationStatementSyntax node)
			{
				var modifiers = ConvertModifiers(node.Modifiers, TokenContext.Local);

				var declarations = new List<LocalDeclarationStatementSyntax>();

				foreach (var declarator in node.Declarators)
					foreach (var decl in SplitVariableDeclarations(declarator, nodesVisitor, semanticModel))
						declarations.Add(SyntaxFactory.LocalDeclarationStatement(modifiers, decl.Value).WithConvertedTriviaFrom(node));

				return SyntaxFactory.List<StatementSyntax>(declarations);
			}

			public override SyntaxList<StatementSyntax> VisitAddRemoveHandlerStatement(VBSyntax.AddRemoveHandlerStatementSyntax node)
			{
				var syntaxKind = node.Kind() == VBasic.SyntaxKind.AddHandlerStatement ? SyntaxKind.AddAssignmentExpression : SyntaxKind.SubtractAssignmentExpression;
				return SingleStatement(SyntaxFactory.AssignmentExpression(syntaxKind,
					(ExpressionSyntax) node.EventExpression.Accept(nodesVisitor),
					(ExpressionSyntax) node.DelegateExpression.Accept(nodesVisitor))
					.WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitExpressionStatement(VBSyntax.ExpressionStatementSyntax node)
			{
				return SingleStatement((ExpressionSyntax)node.Expression.Accept(nodesVisitor).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitAssignmentStatement(VBSyntax.AssignmentStatementSyntax node)
			{
				var kind = ConvertToken(node.Kind(), TokenContext.Local);
				return SingleStatement(SyntaxFactory.AssignmentExpression(kind, (ExpressionSyntax)node.Left.Accept(nodesVisitor), (ExpressionSyntax)node.Right.Accept(nodesVisitor))
					.WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitThrowStatement(VBSyntax.ThrowStatementSyntax node)
			{
				return SingleStatement(SyntaxFactory.ThrowStatement((ExpressionSyntax)node.Expression?.Accept(nodesVisitor)).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitReturnStatement(VBSyntax.ReturnStatementSyntax node)
			{
				if (IsIterator)
					return SingleStatement(SyntaxFactory.YieldStatement(SyntaxKind.YieldBreakStatement).WithConvertedTriviaFrom(node));
				return SingleStatement(SyntaxFactory.ReturnStatement((ExpressionSyntax)node.Expression?.Accept(nodesVisitor)).WithConvertedTriviaFrom(node));
			}

            public override SyntaxList<StatementSyntax> VisitContinueStatement(VBSyntax.ContinueStatementSyntax node)
            {
                return SingleStatement(SyntaxFactory.ContinueStatement().WithConvertedTriviaFrom(node));
            }

            public override SyntaxList<StatementSyntax> VisitYieldStatement(VBSyntax.YieldStatementSyntax node)
			{
				return SingleStatement(SyntaxFactory.YieldStatement(SyntaxKind.YieldReturnStatement, (ExpressionSyntax)node.Expression?.Accept(nodesVisitor)).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitExitStatement(VBSyntax.ExitStatementSyntax node)
			{
				switch (VBasic.VisualBasicExtensions.Kind(node.BlockKeyword))
				{
					case VBasic.SyntaxKind.SubKeyword:
						return SingleStatement(SyntaxFactory.ReturnStatement().WithConvertedTriviaFrom(node));
					case VBasic.SyntaxKind.FunctionKeyword:
						VBasic.VisualBasicSyntaxNode typeContainer = (VBasic.VisualBasicSyntaxNode)node.Ancestors().OfType<VBSyntax.LambdaExpressionSyntax>().FirstOrDefault()
							?? node.Ancestors().OfType<VBSyntax.MethodBlockSyntax>().FirstOrDefault();
						var info = typeContainer.TypeSwitch(
							(VBSyntax.LambdaExpressionSyntax e) => semanticModel.GetTypeInfo(e).Type.GetReturnType(),
							(VBSyntax.MethodBlockSyntax e) => {
								var type = (TypeSyntax)e.SubOrFunctionStatement.AsClause?.Type.Accept(nodesVisitor) ?? SyntaxFactory.ParseTypeName("object");
								return semanticModel.GetSymbolInfo(type).Symbol.GetReturnType();
							}
						);
						ExpressionSyntax expr;
						if (info.IsReferenceType)
							expr = SyntaxFactory.LiteralExpression(SyntaxKind.NullLiteralExpression);
						else if (info.CanBeReferencedByName)
							expr = SyntaxFactory.DefaultExpression(SyntaxFactory.ParseTypeName(info.ToMinimalDisplayString(semanticModel, node.SpanStart)));
						else
							throw new NotSupportedException();
						return SingleStatement(SyntaxFactory.ReturnStatement(expr).WithConvertedTriviaFrom(node));
					default:
						return SingleStatement(SyntaxFactory.BreakStatement().WithConvertedTriviaFrom(node));
				}
			}

            public override SyntaxList<StatementSyntax> VisitRaiseEventStatement(VBSyntax.RaiseEventStatementSyntax node)
            {
                return SingleStatement(
                    SyntaxFactory.ConditionalAccessExpression(
                        (ExpressionSyntax)node.Name.Accept(nodesVisitor),
                        SyntaxFactory.InvocationExpression(
                            SyntaxFactory.MemberBindingExpression(SyntaxFactory.IdentifierName("Invoke")),
                            (ArgumentListSyntax)node.ArgumentList.Accept(nodesVisitor)
                        )
                    ).WithConvertedTriviaFrom(node)
                );
            }

            public override SyntaxList<StatementSyntax> VisitSingleLineIfStatement(VBSyntax.SingleLineIfStatementSyntax node)
			{
				var condition = (ExpressionSyntax)node.Condition.Accept(nodesVisitor);
				var block = SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this)));
				ElseClauseSyntax elseClause = null;

				if (node.ElseClause != null)
				{
					var elseBlock = SyntaxFactory.Block(node.ElseClause.Statements.SelectMany(s => s.Accept(this)));
					elseClause = SyntaxFactory.ElseClause(elseBlock.UnpackBlock());
				}
				return SingleStatement(SyntaxFactory.IfStatement(condition, block.UnpackBlock(), elseClause).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitMultiLineIfBlock(VBSyntax.MultiLineIfBlockSyntax node)
			{
				var condition = (ExpressionSyntax)node.IfStatement.Condition.Accept(nodesVisitor);
				var block = SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this)));
				ElseClauseSyntax elseClause = null;

				if (node.ElseBlock != null)
				{
					var elseBlock = SyntaxFactory.Block(node.ElseBlock.Statements.SelectMany(s => s.Accept(this)));
					elseClause = SyntaxFactory.ElseClause(elseBlock.UnpackBlock());
				}

				foreach (var elseIf in node.ElseIfBlocks.Reverse())
				{
					var elseBlock = SyntaxFactory.Block(elseIf.Statements.SelectMany(s => s.Accept(this)));
					var ifStmt = SyntaxFactory.IfStatement((ExpressionSyntax)elseIf.ElseIfStatement.Condition.Accept(nodesVisitor), elseBlock.UnpackBlock(), elseClause);
					elseClause = SyntaxFactory.ElseClause(ifStmt);
				}

				return SingleStatement(SyntaxFactory.IfStatement(condition, block.UnpackBlock(), elseClause).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitForBlock(VBSyntax.ForBlockSyntax node)
			{
				var stmt = node.ForStatement;
				ExpressionSyntax startValue = (ExpressionSyntax)stmt.FromValue.Accept(nodesVisitor);
				VariableDeclarationSyntax declaration = null;
				ExpressionSyntax id;
				if (stmt.ControlVariable is VBSyntax.VariableDeclaratorSyntax) {
					var v = (VBSyntax.VariableDeclaratorSyntax)stmt.ControlVariable;
					declaration = SplitVariableDeclarations(v, nodesVisitor, semanticModel).Values.Single();
					declaration = declaration.WithVariables(SyntaxFactory.SingletonSeparatedList(declaration.Variables[0].WithInitializer(SyntaxFactory.EqualsValueClause(startValue))));
					id = SyntaxFactory.IdentifierName(declaration.Variables[0].Identifier);
				} else
				{
					id = (ExpressionSyntax)stmt.ControlVariable.Accept(nodesVisitor);
					var symbol = semanticModel.GetSymbolInfo(stmt.ControlVariable).Symbol;
					if (!semanticModel.LookupSymbols(node.FullSpan.Start, name: symbol.Name).Any())
					{
						var variableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(SyntaxFactory.Identifier(symbol.Name), null,
							SyntaxFactory.EqualsValueClause(startValue));
						declaration = SyntaxFactory.VariableDeclaration(
							SyntaxFactory.IdentifierName("var"),
							SyntaxFactory.SingletonSeparatedList(variableDeclaratorSyntax));
					}
					else
					{
						startValue = SyntaxFactory.AssignmentExpression(SyntaxKind.SimpleAssignmentExpression, id, startValue);
					}
				}

				var step = (ExpressionSyntax)stmt.StepClause?.StepValue.Accept(nodesVisitor);
				PrefixUnaryExpressionSyntax value = step.SkipParens() as PrefixUnaryExpressionSyntax;
				ExpressionSyntax condition;
				if (value == null) {
					condition = SyntaxFactory.BinaryExpression(SyntaxKind.LessThanOrEqualExpression, id, (ExpressionSyntax)stmt.ToValue.Accept(nodesVisitor));
				} else {
					condition = SyntaxFactory.BinaryExpression(SyntaxKind.GreaterThanOrEqualExpression, id, (ExpressionSyntax)stmt.ToValue.Accept(nodesVisitor));
				}
				if (step == null)
					step = SyntaxFactory.PostfixUnaryExpression(SyntaxKind.PostIncrementExpression, id);
				else
					step = SyntaxFactory.AssignmentExpression(SyntaxKind.AddAssignmentExpression, id, step);
				var block = SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this)));
				return SingleStatement(SyntaxFactory.ForStatement(
					declaration,
					declaration != null ? SyntaxFactory.SeparatedList<ExpressionSyntax>() : SyntaxFactory.SingletonSeparatedList(startValue),
					condition,
					SyntaxFactory.SingletonSeparatedList(step),
					block.UnpackBlock()).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitForEachBlock(VBSyntax.ForEachBlockSyntax node)
			{
				var stmt = node.ForEachStatement;

				TypeSyntax type = null;
				SyntaxToken id;
				if (stmt.ControlVariable is VBSyntax.VariableDeclaratorSyntax)
				{
					var v = (VBSyntax.VariableDeclaratorSyntax)stmt.ControlVariable;
					var declaration = SplitVariableDeclarations(v, nodesVisitor, semanticModel).Values.Single();
					type = declaration.Type;
					id = declaration.Variables[0].Identifier;
				}
				else
				{
					var v = (IdentifierNameSyntax)stmt.ControlVariable.Accept(nodesVisitor);
					id = v.Identifier;
					type = SyntaxFactory.ParseTypeName("var");
				}

				var block = SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this)));
				return SingleStatement(SyntaxFactory.ForEachStatement(
						type,
						id,
						(ExpressionSyntax)stmt.Expression.Accept(nodesVisitor),
						block.UnpackBlock()
					).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitSelectBlock(VBSyntax.SelectBlockSyntax node)
			{
				var expr = (ExpressionSyntax)node.SelectStatement.Expression.Accept(nodesVisitor);
				SwitchStatementSyntax switchStatement;
				if (ConvertToSwitch(expr, node.CaseBlocks, out switchStatement))
					return SingleStatement(switchStatement.WithConvertedTriviaFrom(node));
				throw new NotSupportedException();
			}

			public override SyntaxList<StatementSyntax> VisitWithBlock(VBSyntax.WithBlockSyntax node)
			{
				var withExpression = (ExpressionSyntax) node.WithStatement.Expression.Accept(nodesVisitor);
				withBlockTempVariableNames.Push(GetUniqueVariableNameInScope(node, "withBlock"));
				try
				{
					var variableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(
						SyntaxFactory.Identifier(withBlockTempVariableNames.Peek()), null,
						SyntaxFactory.EqualsValueClause(withExpression));
					var declaration = SyntaxFactory.LocalDeclarationStatement(SyntaxFactory.VariableDeclaration(
						SyntaxFactory.IdentifierName("var"),
						SyntaxFactory.SingletonSeparatedList(variableDeclaratorSyntax)));
					var statements = node.Statements.SelectMany(s => s.Accept(this));

					return SingleStatement(SyntaxFactory.Block(new[] {declaration}.Concat(statements).ToArray())).WithConvertedTriviaFrom(node));
				}
				finally
				{
					withBlockTempVariableNames.Pop();
				}
			}

			private string GetUniqueVariableNameInScope(SyntaxNode node, string variableNameBase)
			{
				var reservedNames = withBlockTempVariableNames.Concat(node.DescendantNodesAndSelf()
					.SelectMany(syntaxNode => semanticModel.LookupSymbols(syntaxNode.SpanStart).Select(s => s.Name)));
				return NameGenerator.EnsureUniqueness(variableNameBase, reservedNames, true);
			}

			private bool ConvertToSwitch(ExpressionSyntax expr, SyntaxList<VBSyntax.CaseBlockSyntax> caseBlocks, out SwitchStatementSyntax switchStatement)
			{
				switchStatement = null;

				var sections = new List<SwitchSectionSyntax>();
				foreach (var block in caseBlocks)
				{
					var labels = new List<SwitchLabelSyntax>();
					foreach (var c in block.CaseStatement.Cases)
					{
						if (c is VBSyntax.SimpleCaseClauseSyntax) {
							var s = (VBSyntax.SimpleCaseClauseSyntax)c;
							labels.Add(SyntaxFactory.CaseSwitchLabel((ExpressionSyntax)s.Value.Accept(nodesVisitor)));
						} else if (c is VBSyntax.ElseCaseClauseSyntax) {
							labels.Add(SyntaxFactory.DefaultSwitchLabel());
						} else return false;
					}
					var list = SingleStatement(SyntaxFactory.Block(block.Statements.SelectMany(s => s.Accept(this)).Concat(SyntaxFactory.BreakStatement())));
					sections.Add(SyntaxFactory.SwitchSection(SyntaxFactory.List(labels), list));
				}
				switchStatement = SyntaxFactory.SwitchStatement(expr, SyntaxFactory.List(sections));
				return true;
			}

			public override SyntaxList<StatementSyntax> VisitTryBlock(VBSyntax.TryBlockSyntax node)
			{
				var block = SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this)));
				return SingleStatement(
					SyntaxFactory.TryStatement(
						block,
						SyntaxFactory.List(node.CatchBlocks.Select(c => (CatchClauseSyntax)c.Accept(nodesVisitor))),
						(FinallyClauseSyntax)node.FinallyBlock?.Accept(nodesVisitor)
					).WithConvertedTriviaFrom(node)
				);
			}

			public override SyntaxList<StatementSyntax> VisitSyncLockBlock(VBSyntax.SyncLockBlockSyntax node)
			{
				return SingleStatement(SyntaxFactory.LockStatement(
					(ExpressionSyntax)node.SyncLockStatement.Expression.Accept(nodesVisitor),
					SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this))).UnpackBlock()
				).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitUsingBlock(VBSyntax.UsingBlockSyntax node)
			{
				if (node.UsingStatement.Expression == null) {
					StatementSyntax stmt = SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this)));
					foreach (var v in node.UsingStatement.Variables.Reverse())
						foreach (var declaration in SplitVariableDeclarations(v, nodesVisitor, semanticModel).Values.Reverse())
							stmt = SyntaxFactory.UsingStatement(declaration, null, stmt);
					return SingleStatement(stmt.WithConvertedTriviaFrom(node));
				} else {
					var expr = (ExpressionSyntax)node.UsingStatement.Expression.Accept(nodesVisitor);
					return SingleStatement(SyntaxFactory.UsingStatement(null, expr, SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this))).UnpackBlock())
						.WithConvertedTriviaFrom(node));
				}
			}

			public override SyntaxList<StatementSyntax> VisitWhileBlock(VBSyntax.WhileBlockSyntax node)
			{
				return SingleStatement(SyntaxFactory.WhileStatement(
					(ExpressionSyntax)node.WhileStatement.Condition.Accept(nodesVisitor),
					SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this))).UnpackBlock()
				).WithConvertedTriviaFrom(node));
			}

			public override SyntaxList<StatementSyntax> VisitDoLoopBlock(VBSyntax.DoLoopBlockSyntax node)
			{
				if (node.DoStatement.WhileOrUntilClause != null)
				{
					var stmt = node.DoStatement.WhileOrUntilClause;
					if (stmt.WhileOrUntilKeyword.IsKind(VBasic.SyntaxKind.WhileKeyword))
						return SingleStatement(SyntaxFactory.WhileStatement(
							(ExpressionSyntax)stmt.Condition.Accept(nodesVisitor),
							SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this))).UnpackBlock()
						).WithConvertedTriviaFrom(node));
					else
						return SingleStatement(SyntaxFactory.WhileStatement(
							SyntaxFactory.PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, (ExpressionSyntax)stmt.Condition.Accept(nodesVisitor)),
							SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this))).UnpackBlock()
						).WithConvertedTriviaFrom(node));
				}
				if (node.LoopStatement.WhileOrUntilClause != null)
				{
					var stmt = node.LoopStatement.WhileOrUntilClause;
					if (stmt.WhileOrUntilKeyword.IsKind(VBasic.SyntaxKind.WhileKeyword))
						return SingleStatement(SyntaxFactory.DoStatement(
							SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this))).UnpackBlock(),
							(ExpressionSyntax)stmt.Condition.Accept(nodesVisitor)
						).WithConvertedTriviaFrom(node));
					else
						return SingleStatement(SyntaxFactory.DoStatement(
							SyntaxFactory.Block(node.Statements.SelectMany(s => s.Accept(this))).UnpackBlock(),
							SyntaxFactory.PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, (ExpressionSyntax)stmt.Condition.Accept(nodesVisitor))
						).WithConvertedTriviaFrom(node));
				}
				throw new NotSupportedException();
			}

			SyntaxList<StatementSyntax> SingleStatement(StatementSyntax statement)
			{
				return SyntaxFactory.SingletonList(statement);
			}

			SyntaxList<StatementSyntax> SingleStatement(ExpressionSyntax expression)
			{
				return SyntaxFactory.SingletonList<StatementSyntax>(SyntaxFactory.ExpressionStatement(expression));
			}
		}
    }

    static class Extensions
    {
        public static StatementSyntax UnpackBlock(this BlockSyntax block)
        {
            return block.Statements.Count == 1 ? block.Statements[0] : block;
        }
    }
}
