class Metric < ActiveRecord::Base

	has_many :recents
	has_many :data, :through => :recent

	def get_datum(company)
	end

	def compare(datum)
		# Return percentile
	end

end