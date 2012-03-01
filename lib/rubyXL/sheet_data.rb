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

    def each
      @data.each do |c|
        c ||= Array.new
        c.each{ |i| yield i }
      end 
    end

  end
end
