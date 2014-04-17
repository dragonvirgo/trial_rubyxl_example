require 'rubyXL'

workbook = RubyXL::Parser.parse("../data/これはテスト＿20140417.xlsx")
workbook.worksheets.each_with_index do |worksheet, index|
  p "シート名#{index}:#{worksheet.sheet_name}"
end

# シート3を取得
worksheet3 = workbook[2]

row = 0
col = 0

# 数値
p "A1:#{worksheet3[row][col].value}" #=> 1

row += 1

# 日付
p "B1:#{worksheet3[row][col].value}" #=> 2014/4/17

row += 1

# 関数の値
p "C1:#{worksheet3[row][col].value}"  #=> 101

row += 1

# 値が存在しない
# todo:rowが存在してcolが内ケースがあり得るのか？
unless worksheet3[row].nil?
  p "D1:#{worksheet3[row][col].value}" #=> nil
end

row = 0
col += 1

# 数値
p "A2:#{worksheet3[row][col].value}" #=> 1.1

row += 1

# 文字列
p "B2:#{worksheet3[row][col].value}" #=> あああ

row += 1

# 関数の値
p "C2:#{worksheet3[row][col].value}"  #=> 10:59
