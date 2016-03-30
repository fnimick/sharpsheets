using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;

class Cell {
  public int? final_value;
  public INestedEquation contents;
  public int value {
    get {
      if (final_value != null) {
        return final_value.Value;
      }
      final_value = this.contents.compute();
      return final_value.Value;
    }
  }
}

interface INestedEquation {
  int compute();
}

class Value : INestedEquation {
  public Value(int val) {
    this.value = val;
  }
  private int value;
  public int compute() {
    return this.value;
  }
}

class Operation : INestedEquation {
  public Operation(INestedEquation left, INestedEquation right, Func<int,int,int> o) {
    this.left = left;
    this.right = right;
    this.oper = o;
  }
  private INestedEquation left;
  private INestedEquation right;
  private Func<int,int,int> oper;

  public int compute() {
    return this.oper(this.left.compute(), this.right.compute());
  }
}

class Operator {
  private static Func<int,int,int> opAdd = (a,b) => a + b;
  private static Func<int,int,int> opSub = (a,b) => a - b;
  private static Func<int,int,int> opMul = (a,b) => a * b;
  private static Func<int,int,int> opDiv = (a,b) => a / b;

  public static readonly Dictionary<string, Func<int,int,int>> lookup
    = new Dictionary<string, Func<int,int,int>>
    {
      {"+", opAdd},
      {"-", opSub},
      {"*", opMul},
      {"/", opDiv}
    };
}

class Ref : INestedEquation {
  public Ref(Cell target) {
    this.target = target;
  }

  private Cell target;

  public int compute() {
    return this.target.value;
  }
}

public class SharpSheets
{
  static public int Main(string[] args)
  {
    if (args.Length != 2) {
      Console.WriteLine("Please run with <input_file> and <output_file> as arguments.");
      return 1;
    }
    var input_file = args[0];
    var output_file = args[1];
    var start_matrix = new List<string[]>();
    var height = 0;
    try {
      using (StreamReader reader = new StreamReader(input_file)) {
        string line;
        while ((line = reader.ReadLine()) != null) {
          var fields = line.Split(',');
          height = (fields.Length > height) ? fields.Length : height;
          start_matrix.Add(fields);
        }
      }
    } catch {
      Console.WriteLine("Error reading input file. Are you sure it exists and is readable?");
      return 1;
    }
    var width = start_matrix.Count;
    var output_matrix = new Cell[width][];
    for (int i = 0; i < width; i++) {
      output_matrix[i] = new Cell[height];
      for (int j = 0; j < height; j++)
        output_matrix[i][j] = new Cell();
    }
    try {
      process_matrix(start_matrix, output_matrix);
    } catch {
      Console.WriteLine("Error processing input file - are you sure it is a valid spreadsheet?");
      return 1;
    }
    var rows = output_matrix.Select(row => string.Join(",", row.Select(cell => cell.value)));
    var contents = string.Join("\n", rows) + "\n";
    try {
      File.WriteAllText(output_file, contents);
    } catch {
      Console.WriteLine("Error writing output file.");
      return 1;
    }
    return 0;
  }

  static void process_matrix(List<string[]> start_matrix, Cell[][] output) {
    int i = 0, j = 0;
    foreach (string[] row in start_matrix) {
      foreach (string cell in row) {
        output[i][j].contents = process_cell(cell, output);
        j++;
      }
      j = 0;
      i++;
    }
  }

  static INestedEquation process_cell(string input, Cell[][] board) {
    var fields = input.Split(' ');
    int number;
    if (fields.Length == 1) {
      if (int.TryParse(input, out number))
        return new Value(number);
      return convert_ref(input, board);
    }
    INestedEquation left, right;
    if (int.TryParse(fields[0], out number))
      left = new Value(number);
    else
      left = convert_ref(fields[0], board);
    if (int.TryParse(fields[1], out number))
      right = new Value(number);
    else
      right = convert_ref(fields[1], board);
    return new Operation(left, right, Operator.lookup[fields[2]]);
  }

  static int convert_col_string(string input) {
    var total = 0;
    var current_base = 1;
    foreach (char c in input.ToCharArray().Reverse()) {
      total += (char.ToUpper(c) - 64) * current_base;
      current_base *= 26;
    }
    return total - 1;
  }

  static Ref convert_ref(string input, Cell[][] board) {
    Match m = Regex.Match(input, @"\b(\w+)(\d+)\b");
    var col = convert_col_string(m.Groups[1].Value);
    var row = int.Parse(m.Groups[2].Value) - 1;
    return new Ref(board[row][col]);
  }
}
