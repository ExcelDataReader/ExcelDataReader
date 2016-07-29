using System;
using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Commands;

namespace NUnit.Framework {
	/// <summary>
	/// A simple ExpectedExceptionAttribute
	/// </summary>
	[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
	public class ExpectedExceptionAttribute : NUnitAttribute, IWrapTestMethod {
		private readonly Type _expectedExceptionType;

		public ExpectedExceptionAttribute(Type type) {
			_expectedExceptionType = type;
		}

		public TestCommand Wrap(TestCommand command) {
			return new ExpectedExceptionCommand(command, _expectedExceptionType);
		}

		private class ExpectedExceptionCommand : DelegatingTestCommand {
			private readonly Type _expectedType;

			public ExpectedExceptionCommand(TestCommand innerCommand, Type expectedType)
				: base(innerCommand) {
				_expectedType = expectedType;
			}

			public override TestResult Execute(TestExecutionContext context) {
				Type caughtType = null;

				try {
					innerCommand.Execute(context);
				} catch (Exception ex) {
					if (ex is NUnitException)
						ex = ex.InnerException;
					caughtType = ex.GetType();
				}

				if (caughtType == _expectedType)
					context.CurrentResult.SetResult(ResultState.Success);
				else if (caughtType != null)
					context.CurrentResult.SetResult(ResultState.Failure,
						string.Format("Expected {0} but got {1}", _expectedType.Name, caughtType.Name));
				else
					context.CurrentResult.SetResult(ResultState.Failure,
						string.Format("Expected {0} but no exception was thrown", _expectedType.Name));

				return context.CurrentResult;
			}
		}
	}
}
