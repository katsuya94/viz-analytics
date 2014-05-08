require 'date'

class Datum < ActiveRecord::Base

	belongs_to :company
	belongs_to :metric
	belongs_to :previous, :class_name => 'Datum'

	def old?
		lifetime = self.metric.lifetime
		Date.today - self.updated_at.to_date > lifetime
	end

end
