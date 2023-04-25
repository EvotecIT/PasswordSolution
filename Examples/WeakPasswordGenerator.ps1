$Months = @(
    # english
    "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
    # polish
    "Styczen", "Luty", "Marzec", "Kwiecien", "Maj", "Czerwiec", "Lipiec", "Sierpien", "Wrzesien", "Pazdziernik", "Listopad", "Grudzien"
    # spanish
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre'
    # german
    "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"
    # russian
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    # french
    'Janvier', 'Fevrier', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Aout', 'Septembre', 'Octobre', 'Novembre', 'Decembre'
) | Sort-Object -Unique
$Numbers = 0..9
$Years = 2020..2023
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