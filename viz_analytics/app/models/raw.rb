class Raw < Metric

	def get_datum(company)
		recent = self.recents.with_company company
		
		unless recent.datum.old?
			return recent.datum.value, self.class.name.underscore.to_sym
		end

		value = self.class.raw(company)

		new_datum = Datum.new :value => value, :company => company, :previous => recent.datum
		new_datum.save

		recent.datum = new_datum
		recent.save

		return value, self.class.name.underscore.to_sym
	end

	def self.raw(company)
	end

end