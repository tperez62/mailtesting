require 'net/imap'
require 'csv'
require 'mail'
require 'securerandom'
require 'logger'

@server = 'imap.gmail.com'
@user = ''
@pass = ''
@folder = 'INBOX'
# @app_path = '/home/tony/ruby/mailtesting'
@currentDate = Time.now.strftime('%d-%b-%Y')
@port = 993
@use_ssl = true
@attachment_path = 'Invoice Downloads'
@log_path = 'logs'
@downloaded_files = 0
VALID_ATTACHMENT_EXTS = [
	'.pdf'
]

#Creates logger and log path
def setup_logger
	if !Dir.exists?(@log_path)
		Dir.mkdir(@log_path)
	end
	@logger = Logger.new("#{@log_path}/log.log")
end

#Loads CSV with given name, and if it doesn't exist, creates a default csv
def load_csv(csvname)
	begin
		@csvname = csvname
		if !File.exists?(csvname)
			default_csv(csvname)
		end
		csv = CSV.read(csvname, headers:true, converters: :numeric)
		puts "load success"
		@lastSeqno = csv[0]['seqno']
		@lastDate = csv[0]['date']
	rescue => e
		@logger.error(e.message)
		raise e.message
	end
end

#Creates default CSV with seqno=0, and date=currentDate, 3 days back
def default_csv(csvname)
	CSV.open(csvname, "wb") do |csv|
		csv << ['seqno', 'date']
		csv << [0, Date.parse(@currentDate).prev_day(3).strftime('%d-%b-%Y')]
	end
	@logger.info("No CSV, Deafult Created")
end

#Downloads emails and stores in an array of arrays, [email, seqno]
def download_emails
	begin
		emails = []
		imap = Net::IMAP.new @server, port:@port, ssl:@use_ssl
		imap.login @user, @pass
		imap.select @folder
		imap.search(["SINCE", "#{@lastDate}"]).each do |seqno|
			next if @lastSeqno >= seqno
			puts seqno
			body = imap.fetch(seqno, 'RFC822')[0].attr['RFC822']
			email = Mail.new(body)
			emails << [email, seqno]
		end
		imap.logout
		imap.disconnect
		emails
	rescue => e
		raise e.message
	end
end

#Saves valid attachments into a dated directory
def save_attachments
	begin
		if !Dir.exists?(@attachment_path)
			Dir.mkdir(@attachment_path)
			@logger.info("Attachment Path Created")
		end
		working_path = "#{@attachment_path}/#{@currentDate}/"
		if !Dir.exists?(working_path)
			Dir.mkdir(working_path)
			@logger.info("Working Path Created")
		end
		emails = download_emails
		emails.each do |email_arr|
			email_arr[0].attachments.each do |attachment|
				file_ext = File.extname(attachment.filename.downcase)
				
				if VALID_ATTACHMENT_EXTS.include?(file_ext)
					data = attachment.body.decoded
				
					#Creates a random file name and checks for if it already exists
					save_filename = create_random_filename(file_ext)
					
					valid = false
					while (valid == false)
						if File.exists?("#{working_path}#{save_filename}")
							save_filename = create_random_filename(file_ext)
						else
							valid = true
						end
					end
					
					#if File.exists?("#{working_path}#{save_filename}")
						#puts "File already exists"
						#next
					#end
					
					File.open(File.join(working_path, save_filename), "wb") { |f| 
						f.write data
					}
					@downloaded_files += 1
				end
			end
			@currentSeqno = email_arr[1]
		end
		if (@currentSeqno.nil?) 
			@currentSeqno = @lastSeqno
			@logger.info("No Files Downloaded")
		else
			@logger.info("Downloads Successful. Downloaded #{@downloaded_files} files. Last Seqno downloaded was #{@currentSeqno}")
		end
		update_csv
	rescue => e
		@logger.error(e.message)
		raise e.message
	end
end

#Creates a random filename from the current date and 6 random characters
def create_random_filename(file_ext)
	random_str = SecureRandom.hex[0..5]
	save_filename = "#{@currentDate}-#{random_str}#{file_ext}"
	save_filename
end

#Updates the CSV with the last downloaded Seqno and the last date that the script was run
def update_csv
	CSV.open(@csvname, "wb") do |csv|
		csv << ['seqno', 'date']
		csv << [@currentSeqno, @currentDate]
	end
end

setup_logger
@logger.info("Application Executed")
load_csv('testcsv.csv')
save_attachments
