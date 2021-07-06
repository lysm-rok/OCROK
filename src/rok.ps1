#


#GLOBAL VARIABLES : set them with your environnement!
$WorkingDir = "$HOME\"
$output=$WorkingDir+"result.csv"
$database=".database.rok"
$day=Get-Date -Format "MM-dd-yyyy"

###################Functions#####################

# OCRanalysis : use the tesseract binary to permofr OCR analysis of a picture in another process. Wait 2 sec at the end to avoid output file to be lock.
# in : full path of a picture to be analysed
# out : full path of a text file with the analysed data
# psm_type : [1 to 12] OCR configuration
function OCRanalysis ( $in, $out, $psm_type)
{
  $args = $in + " "+ $out + " --psm "+ $psm_type
  start-process -FilePath $NomExe -ArgumentList $args -workingdirectory $WorkingDir
  #Wait 2sec, until OCR finish to process the picture"
  Start-Sleep -s 2
}

# ParseGovernorInfo : parse a OCR file restult, looking for governors info
# tmpfile : text file to parse
# picture_rpath : picture been analysed, used again in case parsing fails
# half_auto : not used anymore?
# try_nb : if parsing fails during first attempt, try a new OCR methode on the picture. if parsing fails on second attempt, show a GUI asking the user for manual help.
# output : cvs strings with governors data
function ParseGovernorInfo($tmpFile, $picture_path, $half_auto, $try_nb)
{
    $line_power =get-content ($WorkingDir + $tmpFile + ".txt") | Where-Object {$_ -like ‘*Power:*’}
    $line_Highest_Power =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Highest*Power*’}
    $line_dead =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Dead*’}
    $line_scout =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Scout*Times*’}
    $line_rss_gathered =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Resources*Gathered*’}
    $line_rss_assistance =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Resource*Assistance*’}
    $line_alliance_help =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Alliance*Help*Times*’}
    $line_victory =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Victory*’}
    $line_defeat =get-content ($WorkingDir + $tmpFile+ ".txt") | Where-Object {$_ -like ‘*Defeat*’}

    #1st line is complexe, OCR gathered name,power and kill point
    $useless=$line_power -match '.*Power'
    $parsed_name =""
    
    try
    {
        $name= $matches[0].substring(0,$matches[0].length-6)
        $parsed_name = $name
    }
    catch
    {
        if ($try_nb -eq "1")
        {
            throw "error while analysing name"
        }
        else
        {
            $name = ""
        }
    }

    #check if the name is already known
    $name = get_name_from_database $name
    #ask to manually check / clean the name before to pursue
    $name = manuel_writing $picture_path $name

    #Once validated, add name in a database to avoid to ask for it during next run of the program
    add_in_database $parsed_name  $name

    Write-Host -ForegroundColor green "name :" $name
    $power_words = $line_power.Split(' ')
    $power= "unknown"
    $kill_points= "unknown"
    For($i=0;$i -lt 10;$i++) 
    { 
   
       if ($power_words[$i] -match "Kill"){
       Write-Host -ForegroundColor green "Power: "$power_words[$i-1] 
       $power=$power_words[$i-1] 
       }

       if ($power_words[$i] -match "Points"){
       Write-Host -ForegroundColor green "Kill Points: "$power_words[$i+1]
       $kill_points= $power_words[$i+1]
       }
    }

    $highest_power = $line_Highest_Power.Split(' ')
    $dead =$line_dead.Split(' ')
    $scout = $line_scout.Split(' ')
    $rss_gathered=$line_rss_gathered.Split(' ')
    $rss_assistance =$line_rss_assistance.Split(' ')

    $alliance_help=$line_alliance_help.Split(' ')
    $victory=$line_victory.Split(' ')
    $defeat=$line_defeat.Split(' ')

    Write-Host -ForegroundColor green "Highest power :" $highest_power[2]
    Write-Host -ForegroundColor green "Deads :" $dead[1]
    Write-Host -ForegroundColor green "Resources gathered :" $rss_gathered[2]
    Write-Host -ForegroundColor green "Resources assistance :" $rss_assistance[2]
    Write-Host -ForegroundColor green "Scout :" $scout[2]
    Write-Host -ForegroundColor green "Alliance help :" $alliance_help[3]
    Write-Host -ForegroundColor green "Victory :" $victory[1]
    Write-Host -ForegroundColor green "Defeat :" $defeat[1]
    Write-Host 
    $line_csv= $day+";"+$name +";"+ $power +";"+ $highest_power[2]+";"+$kill_points +";"+ $dead[1] +";"+ $rss_gathered[2]+";"+$rss_assistance[2]+";"+$scout[2]+";"+$alliance_help[3]+";"+$victory[1]+";"+$defeat[1]
    return $line_csv
}

function get_name_from_database($parsed_name)
{
    if (Test-Path $WorkingDir$database)
    {
          $regex= [regex]::escape($parsed_name)
          
          if (get-content ($WorkingDir+$database) | Where-Object {$_ -like "*"+$regex+'*'})
          {
              $name_line =get-content ($WorkingDir+$database) | Where-Object {$_ -like "*"+$regex+'*'}
              $name_in_db=$name_line.Split(";")[1]
              return $name_in_db
          }
          else
          {
           return $parsed_name
           }
    }
    else
    {
        return $parsed_name
    }
    
}

# add_in_database : add a pair {parsed name, validated name} into a database file; it will make the process faster during next run of the program
# parsed_name : a name retrieved by the parser, which may be false
# name : the name validated or written by the user
function add_in_database ($parsed_name , $name)
{
    $line_reg = [regex]::escape($parsed_name) +";"+ $name

    if (!(Test-Path $WorkingDir$database))
    {
       New-Item -path $WorkingDir -name $database -type "file"
       Add-Content -Encoding UTF8 -path $WorkingDir$database -value $line
    }
    else
    {
      if (!(get-content ($WorkingDir+$database) | Where-Object {$_ -like $parsed_name+‘;*’}))
      {
          Add-Content -Encoding UTF8 -path $WorkingDir$database -value $line
      }
    }
}

function Select-File($message='Selectionner un répertoire', $path = 0)
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Binary (*.exe)|*.exe'
    Title = 'Please find and select tessereact.exe file'
    }
    $result  = $FileBrowser.ShowDialog()
    return $FileBrowser.FileName
}

function Select-Folder($message='Selectionner un répertoire', $path = 0)
{
    $object = New-Object -comObject Shell.Application
 
    $folder = $object.BrowseForFolder(0, $message, 0, $path)
    if ($folder -ne $null)
    {
     $folder.self.Path
    }
}

function manuel_writing($picture_path , $name)
{
    $picture = (get-item $picture_path)
    $img = [System.Drawing.Image]::Fromfile($picture);

    $form = New-Object System.Windows.Forms.Form
    $form.Text = '1164 magic tricks'
    $form.Size = New-Object System.Drawing.Size($img.Width,$img.Height)
    $form.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point((($img.Width/2)-350),(($img.Height/2) - 200))
    $label.Size = New-Object System.Drawing.Size(700,80)
    $label.Text = "Magic has limits and it may not worked on governor's names. Please check and correct:"
    $label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,[System.Drawing.FontStyle]::Regular)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point((($img.Width/2)-350),(($img.Height/2) - 100))
    $textBox.Size = New-Object System.Drawing.Size(700,50)
    $textBox.Text =$name
    $textBox.Font =[System.Drawing.Font]::new('Microsoft Sans Serif', 15, [System.Drawing.FontStyle]::Regular)
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point((($img.Width/2)-75),($img.Height/2))
    $okButton.Size = New-Object System.Drawing.Size(150,50)
    $okButton.Font =[System.Drawing.Font]::new('Microsoft Sans Serif', 15, [System.Drawing.FontStyle]::Regular)
    $okButton.Text = 'DONE'

    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $image1 = New-Object System.Windows.Forms.pictureBox
    $image1.Location = New-Object System.Drawing.Size(0,1)
    $image1.Size = New-Object System.Drawing.Size($img.Width,$img.Height)
    $image1.Image = $img

    $form.controls.add($image1)
    $form.Topmost = $true
    $form.Add_Shown({$textBox.Select()})
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $x = $textBox.Text
       # write-host $x
        return $x
    }
}

function print_banner()
{
write-host @"
 .d88888b.   .d8888b.  8888888b.   .d88888b.  888    d8P  
d88P" "Y88b d88P  Y88b 888   Y88b d88P" "Y88b 888   d8P   
888     888 888    888 888    888 888     888 888  d8P    
888     888 888        888   d88P 888     888 888d88K     
888     888 888        8888888P"  888     888 8888888b    
888     888 888    888 888 T88b   888     888 888  Y88b   
Y88b. .d88P Y88b  d88P 888  T88b  Y88b. .d88P 888   Y88b  
 "Y88888P"   "Y8888P"  888   T88b  "Y88888P"  888    Y88b 
                                                          
                                                                                                                    
"@
write-host "                                 Build by the 1164 kingdom"
write-host
write-host
write-host "You could report issues here :https://github.com/lysm-rok/OCROK/issues"
write-host "However this is a one time tool, I will probably don't touch it anymore"
write-host
write-host
}


############################Mains#############################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$half_auto= $true
print_banner

if (Test-Path "c:\Program Files\Tesseract-OCR\tesseract.exe")
{
    $NomExe = "c:\Program Files\Tesseract-OCR\tesseract.exe"
}
elseif (Test-Path "c:\Program Files (x86)\Tesseract-OCR\tesseract.exe")
{
    $NomExe = "c:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
}
else
{
    $NomExe = Select-File "If you have already downloaded and installed tesseract, please give us its path. If not, please check-out the requirements." "c:\"
    if($NomExe -match '.*tesseract.exe')
    {
        Write-host "Tesseract found" $NomExe
    }
    else
    {
    Write-host "tesseract not found, please install it from https://github.com/UB-Mannheim/tesseract/wiki "
    exit
    }
}


$screenDir =Select-Folder 'Please choose the folder with your screenshots'
if (-not([STRING]::IsNullOrEmpty($screenDIr)) -and (Test-Path $screenDir))
{
    try
    {
        Get-ChildItem -File -Path $screenDir| Rename-Item -NewName { $_.Name -replace ' ','' }
    }
    catch
    {
        Write-error "error while checking file names. Close all programm using the pictures and run again this analysis"
    }
}
else
{
    exit
}


Add-Content -Path $output -Value "Date;Name;Power;Highest power;Kill Points;Deads;Resources gathered;Resources assistance;Scout;Alliance Help;Victory;Defeat"
$MonFolder = Get-ChildItem -Path $screenDir -File #On récupère la liste des fichiers de ce répertoire
foreach ($MyFile in $MonFolder)
{

    $imgName=$($MyFile.name)
    $time=Get-Date -Format "MM-dd-yyyy_HH_mm_ss"
    $tmpFile= $time+$imgName.Split('.')[0]
    
    #OCR executable must be installed first
    OCRanalysis $MyFile.FullName ($WorkingDir +$tmpFile) 3

    $tmpFilewithPath =($WorkingDir + $tmpFile + ".txt")
    
    try
    {
        $line_csv=""
        $try_nb="1"
        $line_csv=ParseGovernorInfo $tmpFile $MyFile.FullName $half_auto $try_nb
        Remove-Item $tmpFilewithPath
        Add-Content -Path $output -Value $line_csv
    }
    catch
    {
        $line_csv=""
        Remove-Item $tmpFilewithPath
        $line_csv=""
        OCRanalysis $MyFile.FullName ($WorkingDir +$tmpFile) 6

        $tmpFilewithPath =($WorkingDir + $tmpFile + ".txt")

        $try_nb="2"
        $line_csv= ParseGovernorInfo $tmpFile $MyFile.FullName $half_auto $try_nb
        Remove-Item $tmpFilewithPath
        Add-Content -Path $output -Value $line_csv
    }
}
    write-host "The result file has been created here:" $output

Import-Csv -Path $output -Delimiter ";" -Header date, name, power, "highest power", "kill points", deads, "rss gathered", "rss assistance", scout, "alliance help", victory, defeat | Select-Object -Skip 1 |Out-GridView –Title "1164-tools $output"


