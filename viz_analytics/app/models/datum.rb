require 'date'

class Datum < ActiveRecord::Base

	belongs_to :company
	belongs_to :metric
	has_one :previous, :class_name => 'Datum'

	def old?
		lifetime = self.metric.lifetime
		if Date.today - self.updated_at > lifetime
			return true
		else
			return false
		end
	end

end
