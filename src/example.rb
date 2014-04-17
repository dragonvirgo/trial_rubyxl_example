require 'rubyXL'
require 'benchmark'

workbook = RubyXL::Parser.parse("../data/これはテスト＿20140417.xlsx")
workbook.worksheets.each_with_index do |worksheet, index|
  p "シート名#{index}:#{worksheet.sheet_name}"
end

p "--------------------"

worksheet1 = workbook[0]

puts Benchmark::CAPTION
puts Benchmark.measure{

  col = 0
  (0..2499).each do |row|
    p "cell(#{row}, #{col}):#{worksheet1[row][col].value}"
  end
}

p "--------------------"

worksheet1 = workbook[0]

puts Benchmark::CAPTION
puts Benchmark.measure{

  col = 1
  val = 0
  (0..2499).each do |row|
    val += worksheet1[row][col].value
  end

  p "sum(1〜2500) = #{val}"
}

p "--------------------"

worksheet2 = workbook[1]

puts Benchmark::CAPTION
puts Benchmark.measure{
  (0..49).each do |row|
    (0..51).each do |col|
      worksheet2[row][col].value
      # p "cell(#{row}, #{col}):#{worksheet2[row][col].value}"
    end
  end
}

p "--------------------"

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


