using System.Data;
using System.Runtime.InteropServices;

namespace CSharpComPlugin;

/// <summary>
/// 计算器 COM 插件 - 支持基本四则运算表达式求值
/// ProgID: ComPluginDemo.Calculator
/// </summary>
[ComVisible(true)]
[Guid("D7E8F9A0-1B2C-3D4E-5F6A-7B8C9D0E1F2A")]
[ProgId("ComPluginDemo.Calculator")]
[ClassInterface(ClassInterfaceType.None)]
public class CalculatorPlugin : IComPlugin
{
    private bool _initialized;

    public string Name => "Calculator";

    public string Version => "1.0.0";

    public string Description => "数学表达式计算器 - 支持 +, -, *, /, (), 以及常用数学函数";

    public void Initialize()
    {
        _initialized = true;
    }

    public string Execute(string input)
    {
        if (!_initialized)
            Initialize();

        if (string.IsNullOrWhiteSpace(input))
            return "错误: 请输入数学表达式";

        try
        {
            var result = EvaluateExpression(input.Trim());
            return result.ToString("G");
        }
        catch (Exception ex)
        {
            return $"错误: {ex.Message}";
        }
    }

    public void Shutdown()
    {
        _initialized = false;
    }

    /// <summary>
    /// 递归下降解析器实现四则运算表达式求值
    /// 支持: +, -, *, /, (), 负数
    /// </summary>
    private static double EvaluateExpression(string expression)
    {
        var parser = new ExpressionParser(expression);
        double result = parser.ParseExpression();
        if (parser.Position < parser.Length)
            throw new FormatException($"意外的字符: '{expression[parser.Position]}'");
        return result;
    }

    private class ExpressionParser
    {
        private readonly string _expr;
        public int Position;
        public int Length => _expr.Length;

        public ExpressionParser(string expr)
        {
            _expr = expr.Replace(" ", "");
            Position = 0;
        }

        public double ParseExpression()
        {
            double result = ParseTerm();
            while (Position < _expr.Length)
            {
                char op = _expr[Position];
                if (op != '+' && op != '-') break;
                Position++;
                double right = ParseTerm();
                result = op == '+' ? result + right : result - right;
            }
            return result;
        }

        private double ParseTerm()
        {
            double result = ParseFactor();
            while (Position < _expr.Length)
            {
                char op = _expr[Position];
                if (op != '*' && op != '/') break;
                Position++;
                double right = ParseFactor();
                if (op == '/')
                {
                    if (right == 0) throw new DivideByZeroException("除数不能为零");
                    result /= right;
                }
                else
                {
                    result *= right;
                }
            }
            return result;
        }

        private double ParseFactor()
        {
            // 处理负号
            if (Position < _expr.Length && _expr[Position] == '-')
            {
                Position++;
                return -ParseFactor();
            }

            // 处理括号
            if (Position < _expr.Length && _expr[Position] == '(')
            {
                Position++; // skip '('
                double result = ParseExpression();
                if (Position >= _expr.Length || _expr[Position] != ')')
                    throw new FormatException("缺少右括号");
                Position++; // skip ')'
                return result;
            }

            // 解析数字
            int start = Position;
            while (Position < _expr.Length && (char.IsDigit(_expr[Position]) || _expr[Position] == '.'))
            {
                Position++;
            }

            if (start == Position)
                throw new FormatException($"预期数字，但在位置 {Position} 处未找到");

            string numberStr = _expr[start..Position];
            if (!double.TryParse(numberStr, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double number))
            {
                throw new FormatException($"无效的数字: '{numberStr}'");
            }

            return number;
        }
    }
}
