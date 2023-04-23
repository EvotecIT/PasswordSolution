$Months = @(
    "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
    "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec", "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"
    "Styczen", "Luty", "Marzec", "Kwiecien", "Maj", "Czerwiec", "Lipiec", "Sierpien", "Wrzesien", "Pazdziernik", "Listopad", "Grudzien"
    "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
) | Sort-Object -Unique
$Numbers = 0..9
$Years = 2019..2023
$SpecialChar = @("!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "[", "]", "{", "}", "|", "\")

$Passwords = foreach ($Year in $Years) {
    Write-Color -Text "Year: ", $Year -Color Yellow, White
    $YearPasswords = foreach ($month in $months) {
        foreach ($number in $numbers) {
            foreach ($special in $SpecialChar) {
                $month + $Year.ToString() + $number.ToString() + $special
                $Year.ToString() + $month + $number.ToString() + $special
                $month + $Year.ToString() + $special
            }
        }
    }
    Write-Color -Text "Year: ", $Year, " passwords created: ", $YearPasswords.Count -Color Yellow, White
    $YearPasswords
}
$Passwords.Count