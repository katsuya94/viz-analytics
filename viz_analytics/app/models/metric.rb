class Metric < ActiveRecord::Base

	has_many :data
	has_many :recents

	def get_datum(company)
	end

end

class Raw < Metric

	def get_datum(company)
		last = self.recents.with_company company
		
		unless last.datum.old?
			return self.class.name.underscore.to_sym, last.datum.value
		end

		value = self.class.raw(company)

		new_datum = self.data.build { :value => value, :company => company, :previous => last.datum }
		last.datum = new_datum

		return self.class.name.underscore.to_sym, value
	end

	def self.raw(company)
	end

end

class Calculated < Metric

	class_attribute :dependencies
	self.dependencies = Array.new

	def self.dependency(classname)
		self.dependencies.push(classname.to_s.classify.constantize)
	end

	def get_datum(company)
		last = self.recents.with_company company

		unless last.datum.old?
			return self.class.name.underscore.to_sym, last.datum.value
		end

		sources = Hash.new

		self.dependencies.each do |d|
			key, val = d.get_datum(business)
			sources[key] = val
		end
		
		value = self.class.calculate(sources)

		new_datum = self.data.build { :value => value, :company => company, :previous => last.datum }
		last.datum = new_datum

		return self.class.name.underscore.to_sym, value
	end

	def self.calculate(sources)
	end

end