module RubyXL
  class SheetData
    include Enumerable

    def initialize(values)
      @data = values || Array.new
    end
  
    def [](x)
      @data[x] ||= Array.new
      @data[x]
    end
  
    def []=(x,value)
      @data[x] ||= Array.new
      @data[x] = value
    end

  end
end
