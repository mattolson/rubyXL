module RubyXL
  class SheetData
    include Enumerable

    def initialize(x=1,y=1)
      @data = Array.new(x) {Array.new(y)}
    end
  
    def [](x, y=nil)
      @data[x] ||= Array.new
      
      if y
        @data[x][y]
      else
        @data[x]
      end
    end
  
    def []=(x, y, value)
      @data[x] ||= Array.new
      @data[x][y] = value
    end

    def each
      @data.each do |row|
        row ||= Array.new
        yield row
      end 
    end
    
    def clear_row(x)
      @data[x] = nil
    end

  end
end
