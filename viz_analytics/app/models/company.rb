class Company < ActiveRecord::Base

	has_many :data
	has_many :recents

end
