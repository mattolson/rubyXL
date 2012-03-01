module RubyXL
class SheetData
  include Enumerable

  def initialize(values)
    @data = values || Array.new
  end
  
  def [](x,y)
    @data[x] ||= Array.new
    @data[x][y]
  end
  
  def []=(x,y,value)
    @data[x] ||= Array.new
    @data[x][y] = value
  end
end
