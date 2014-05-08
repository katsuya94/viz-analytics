class Recent < ActiveRecord::Base

	belongs_to :metric
	belongs_to :datum

	scope :with_company, ->(company) { joins(:datum).where( :data => { :company_id => company.id } ).first }

end
