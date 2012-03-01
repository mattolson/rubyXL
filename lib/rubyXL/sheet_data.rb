module RubyXL
  class SheetData
    include Enumerable

    def initialize(values)
      @data = values || Array.new
    end
  
    def [](x, y=nil)
      @data[x] ||= Array.new
      
      if y
        @data[x][y]
      else
        @data[x]
      end
    end
  
    def []=(x,y,value)
      @data[x] ||= Array.new
      @data[x][y] = value
    end

    def each
      @data.each do |c|
        c.each{ |i| yield i }
      end 
    end

  end
end
