class Recent < ActiveRecord::Base

	has_one :datum
	belongs_to :metric

	scope :with_company, ->(company) { joins(:data).where(:company => company) }

end
