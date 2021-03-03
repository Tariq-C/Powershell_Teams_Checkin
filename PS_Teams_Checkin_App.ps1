# Author: Tariq Chatur
# Purpose: 
# A simple check in application that updates a Channel with Check in Information
# Automatically grabs user's full display name
# Uses Team's incoming webhook

# Team's Channel Webhook Uri 
# TODO: Uncomment and add your own URI
$inWebhookUri = ""

# Function Name: checking_in_form
# Parameters: None
# Returns: Status to be published in Teams

# 0 = Checking In
# 1 = Out for Lunch
# 2 = Back From Lunch
# 3 = Checking Out

Function checking_in_form(){
	# Step 1: Definte the form
	#___________________________________________________________________________________________________________________________

	#
	#   Build the form that the user will use to genereate emails
	#

  
  $names = "Checking In", "Out to Lunch", "Back From Lunch" ,"Checking Out"
  $count = $names.Length

  # An array of radio buttons. One for each email template
  $radioButtons = [System.Object[]]::new($count)

  # Initialisation of the button sizes so they remain constant
  $bw = 250
  $bh = 30
  $bi = 10

  # A counter to keep track of current button
  $i = 0

  # return value which will indicate which button was selected
  $x = -1

  # Start of the form Creation, Initialisation of the form
	$Form = New-Object System.Windows.Forms.Form
	$Form.Text = 'Choose an Email Template'
	$Form.Width = 800;
	$Form.Height = 150 + ($bi + $bh)*$count;

	# Set the font of the text to be used within the form
	$Font = New-Object System.Drawing.Font("Calibri",11.5)
	$Form.Font = $Font

  # A loop that creates a radio button and labels them based on the values found in the html files
  foreach ($button in $radioButtons){
    $name = $names[$i]
    $radioButtons[$i] = New-Object System.Windows.Forms.RadioButton
    $radioButtons[$i].Left = $Form.Width - $bi - $bw;
    $radioButtons[$i].Top = $bi * ($i+1) + $bh * $i;
    $radioButtons[$i].Height = $bh;
    $radioButtons[$i].Width = $bw;
    $radioButtons[$i].Text = $name;
    $radioButtons[$i].Checked = $false
    $Form.Controls.Add($radioButtons[$i])
    $i = $i + 1
  }

  # Create the OK button
  $OKButton = new-object System.Windows.Forms.Button
  $OKButton.Left = $bi + $bw + $bi;
  $OKButton.Top = $bi*2 + $bh;
  $OKButton.Height = $bh;
  $OKButton.Width = $bw;
  $OKButton.Text = 'OK'
  $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
  $Form.Controls.Add($OKButton)


	# Create a Cancel button control
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Left = $bi;
	$CancelButton.Top = $bi*2 + $bh;
	$CancelButton.Height = $bh;
	$CancelButton.Width = $bw;
	$CancelButton.Text = 'Cancel'
	$CancelButton.DialogResult = 'Cancel'
	$Form.Controls.Add($CancelButton) # add the Cancel control to the form

	$Form.Topmost = $True                    # set the form to the foreground

	# Step 2: Open the form
	#___________________________________________________________________________________________________________________________

	$Form.Add_Shown({$Form.Activate() })     # activate / display the form

	# Step 3: Form Logic
	#___________________________________________________________________________________________________________________________

  # If Cancel is clicked then exit the form
	if('Cancel' -eq $Form.ShowDialog()){
	    # pause
	    Exit
	}else{

      # Loop through the buttons and see which one is checked then save it to be x
      $j =  0
      foreach($button in $radioButtons){
        if($radioButtons[$j].Checked){
          $x = $j
        }
        $j++
      }

			# return the ticket number and the email selection
			return $names[$x]
    }
}

# Function Name: program_loop
# Parameters: $inWebhookUri - The address to have the status posted to
# Returns: None

Function program_loop($inWebhookUri){
    # Run Boolean 
    $run = $True 
    
    # Application Loop
    while ($run -eq $True){

      # Store the body of the message + Prefix
      $body = ""
      $body +=  "{`"text`":`""

      # Get the status that the user would like to state
      $status = checking_in_form


      # Obtain the Windows Display Name 
      $dom = $env:userdomain
      $usr = $env:username
      $person = ([adsi]"WinNT://$dom/$usr,user").fullname

    
      # Convert to Spoken Word Format
      $body +=  $person + "is" + $status

      # Add Suffix
      $body += "`"}"


      # Post to Teams Channel via Incoming Webhook
      Invoke-RestMethod -Method post -ContentType 'Application/JSON' -Body $body -uri $inWebhookUri
    }

}

program_loop($inWebhookUri)