class Calculated < Metric

	class_attribute :dependencies
	self.dependencies = Array.new

	def self.dependency(classname)
		self.dependencies.push(classname.to_s.classify.constantize)
	end

	def get_datum(company)
		last = self.recents.find_by_company(company)

		unless last.datum.old?
			return self.class.name.underscore.to_sym, last.datum.value
		end

		sources = Hash.new

		self.dependencies.each do |d|
			key, val = d.get_datum(business)
			sources[key] = val
		end
		
		value = self.class.calculate(sources)

		new_datum = Datum.new( :value => value, :company => company, :previous => last.datum )
		self.recent = new_datum

		return value, self.class.name.underscore.to_sym
	end

	def self.calculate(sources)
	end

end