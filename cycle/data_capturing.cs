using IronXL;
using Microsoft.VisualStudio.TestPlatform.Utilities;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Xml.Linq;


namespace cycle
{


    public class data_capturing
    {
        public IWebDriver driver;
        Include include = new Include();



        [Fact]
        public void Test1()
        {

            this.cycle();

        }

        public void cycle()
        {


            WorkBook workbook = WorkBook.Load("D:\\Selenium\\cycle\\DataSheet.xlsx");
            WorkSheet worksheet = workbook.GetWorkSheet("Wedding");
            var range = worksheet["B2:B20"];

            include.Cromelink();
            include.Sleep();
            include.Xpaths("/html/body/app-root/div/div[1]/iw-product-select/div/div/div[8]/div/div[2]");
            include.Sleep(); include.Sleep();
            include.Xpaths("/html/body/app-root/div/div[1]/iw-cycle-bike/div/form/div[2]/div/button");
            include.Sleep();



            include.Xpaths2("//*[@id=\"codelist-search_makeId\"]","B2", worksheet);
            include.Sleep();

            include.Xpaths("//*[@id=\"question_makeId\"]/iw-question-wrapper/div/div[2]/iw-codelist-search/div/div");
            include.Sleep();

            //include.Xpaths2("//*[@id=\"text-input_model\"]", "Other");
            //include.Sleep();

            ////Is it an electric bike?

            //include.Xpaths("//*[@id=\"bool-input_electric_false_button\"]");
            //include.Sleep();


            //include.Xpaths("//*[@id=\"bool-input_electric_true_button\"]");
            //include.Sleep();


            ////What type of bike is it?
            //include.Xpaths("//*[@id=\"select-input_typeId\"]");

            //include.Xpaths("//*[@id=\"select-input_typeId_option_7\"]");
            //include.Sleep();

            ////Can you confirm that your electric bike power output does not exceed 250 watts?

            //include.Xpaths("//*[@id=\"bool-input_exceed250w_true_button\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"bool-input_exceed250w_false_button\"]");
            //include.Sleep();

            //include.Xpaths("/html/body/app-root/div/div[1]/iw-cycle-bike/iw-modal[1]/div[2]/div[3]/div/div[1]/button");
            //include.Sleep();

            ////Can you confirm that your electric motor will not assist you when you are travelling more than 25km/h (15.5mph)?

            //include.Xpaths("//*[@id=\"bool-input_travelling25km_true_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_travelling25km_false_button\"]");
            //include.Sleep();
            //include.Xpaths("/html/body/app-root/div/div[1]/iw-cycle-bike/iw-modal[2]/div[2]/div[3]/div/div[1]/button");
            //include.Sleep();


            ////Did you purchase the bike brand new or second hand?
            //include.Xpaths("//*[@id=\"bool-input_boughtdNew_false_button\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"bool-input_boughtdNew_true_button\"]");
            //include.Sleep();

            ////When did you buy the bike?
            //include.Xpaths("//*[@id=\"button_purchaseDate\"]");

            ////What is the value of your bike?
            //include.Xpaths2("//*[@id=\"text-input_value\"]", "5000");
            //include.Sleep();

            ////Where do you normally store your bike when not in use?
            //include.Xpaths("//*[@id=\"select-input_storeId\"]");

            //include.Xpaths("//*[@id=\"select-input_storeId_option_1\"]");
            //include.Sleep();

            ////Will you leave the bike locked and unattended away from home?
            //include.Xpaths("//*[@id=\"bool-input_leaveUnattendedAwayHome_false_button\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"bool-input_leaveUnattendedAwayHome_true_button\"]");
            //include.Sleep();


            //include.TakingHTML2CanvasFullPageScreenshot();


            ////save button1
            //include.Xpaths("//*[@id=\"question_bikes\"]/iw-multi-instance-question/div/div/div[2]/button");
            //include.Sleep();

            ////Would you like cover for any accessories?
            //include.Xpaths("//*[@id=\"bool-input_otherAccessories_false_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_otherAccessories_true_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"select-input_cover_accessoriesCoverId\"]");
            //include.Xpaths("//*[@id=\"select-input_cover_accessoriesCoverId_option_250\"]");
            //include.Sleep();


            ////next page
            //include.Xpaths("/html/body/app-root/div/div[1]/iw-cycle-bike/div/form/div[2]/div/button");
            //include.Sleep();
            //include.Sleep();
            ////...............................................................................

            ////What is your title?
            //include.Xpaths("//*[@id=\"select-input_proposer_titleId_button_002\"]/span");
            //include.Sleep();

            ////What is your first name?
            //include.Xpaths2("//*[@id=\"text-input_proposer_forename\"]", "iwonder");
            //include.Sleep();

            ////What is your last name?
            //include.Xpaths2("//*[@id=\"text-input_proposer_surname\"]", "testing");
            //include.Sleep();

            ////What is your date of birth?
            //include.Xpaths2("//*[@id=\"date-select_proposer_dateOfBirth_day_input\"]", "30");
            //include.Xpaths2("//*[@id=\"date-select_proposer_dateOfBirth_month_input\"]", "09");
            //include.Xpaths2("//*[@id=\"date-select_proposer_dateOfBirth_year_input\"]", "1994");
            //include.Sleep();

            ////What is your house number/name (optional)?
            //include.Xpaths2("//*[@id=\"text-input_houseSearch\"]", "3");
            //include.Sleep();

            ////What is your postcode?
            //include.Xpaths2("//*[@id=\"text-input_postCode\"]", "SO21 3FL");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"question_proposer_correspondanceAddress\"]/iw-address/form/iw-question-wrapper[2]/div/div[2]/iw-text-input/div/div[2]/button");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"question_proposer_correspondanceAddress\"]/iw-address/form/div/iw-question-wrapper/div/div[2]/div/div/div[2]");
            //include.Sleep();

            ////next2
            //include.Xpaths("/html/body/app-root/div/div[1]/iw-cycle-policyholder/div/form/div[2]/div[2]/button");
            //include.Sleep();


            ////Cover
            ////Would you like to include cover for travel to and from your place of work? 
            //include.Xpaths("//*[@id=\"bool-input_cover_commutingUse_false_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_cover_commutingUse_true_button\"]");
            //include.Sleep();

            ////Would you like to include cover for competitions?
            //include.Xpaths("//*[@id=\"bool-input_cover_competitionUse_true_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_cover_competitionUse_false_button\"]");
            //include.Sleep();

            ////Would you like to cover for non-competitive cycle events?
            //include.Xpaths("//*[@id=\"bool-input_cover_nonCompetitionEventUse_false_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_cover_nonCompetitionEventUse_true_button\"]");
            //include.Sleep();

            ////Would you like to cover for commercial use?
            //include.Xpaths("//*[@id=\"bool-input_cover_commercialUse_true_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_cover_commercialUse_false_button\"]");
            //include.Sleep();

            ////Would you like UK only, EU or Worldwide cover?
            //include.Xpaths("//*[@id=\"select-input_cover_locationId\"]");
            //include.Xpaths("//*[@id=\"select-input_cover_locationId_option_1\"]");

            ////Would you like to include cover to protect you in the event that you injure a person or someone’s property whilst out riding and you have to pay them compensation?
            //include.Xpaths("//*[@id=\"bool-input_publicLiability_false_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_publicLiability_true_button\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"select-input_cover_publicLiabilityId\"]");
            //include.Xpaths("//*[@id=\"select-input_cover_publicLiabilityId_option_1\"]");

            ////Do you require personal accident cover?
            //include.Xpaths("//*[@id=\"bool-input_accidentCover_true_button\"]");

            //include.Xpaths("//*[@id=\"select-input_personalAccidentCoverId\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"select-input_personalAccidentCoverId_option_10000\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_accidentCover_false_button\"]");

            ////Do you require legal expenses cover?
            //include.Xpaths("//*[@id=\"bool-input_cover_legalExpensesCover_true_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_cover_legalExpensesCover_false_button\"]");

            ////Do you require legal expenses cover?
            //include.Xpaths("//*[@id=\"bool-input_replacementBikeHire_false_button\"]");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"bool-input_replacementBikeHire_true_button\"]");

            //include.Xpaths("//*[@id=\"select-input_replacementBikeHireId\"]");
            //include.Xpaths("//*[@id=\"select-input_replacementBikeHireId_option_500\"]");
            //include.Sleep();

            ////What date would you like cover to start?
            //include.Xpaths("//*[@id=\"select-input_cover_coverStartDate_option_Tue Feb 28 2023 00:00:00 GMT+0000\"]");
            //include.Sleep();

            ////How are you going to pay for your cycle insurance ?
            //include.Xpaths("//*[@id=\"select-input_paymentPreferenceId_button_Monthly\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"select-input_paymentPreferenceId_button_Annually\"]/span");
            //include.Sleep();

            ////Do you require cover for your immediate family member(s) who live at the same address?
            //include.Xpaths("//*[@id=\"bool-input_familyCoverRequired_false_button\"]//*[@id=\"bool-input_familyCoverRequired_false_button\"]");
            //include.Sleep();


            ////What is your family member's first name?
            //include.Sleep();
            //include.Xpaths2("//*[@id=\"text-input_houseSearch\"]", "john");
            //include.Sleep();
            ////What is your family member's last name?
            //include.Xpaths2("//*[@id=\"text-input_surname\"]", "doe");
            //include.Sleep();

            //include.Xpaths2("//*[@id=\"date-select_dateOfBirth_day_input\"]", "01");
            //include.Xpaths2("//*[@id=\"date-select_dateOfBirth_month_input\"]", "01");
            //include.Xpaths2("//*[@id=\"date-select_dateOfBirth_year_input\"]", "1990");

            //include.Xpaths("//*[@id=\"question_riders\"]/iw-multi-instance-question/div/div/div[2]/button");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"question_riders\"]/iw-multi-instance-question/div/div/iw-question-wrapper/div/div[2]/div/div/div[1]/span");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"question_riders\"]/iw-multi-instance-question/div/div/div[1]/button");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"question_riders\"]/iw-multi-instance-question/div/div/iw-question-wrapper/div/div[2]/div/div/div[2]/span");
            //include.Sleep();


            ////How many years No Claims Discount do you have on a specific cycle insurance policy?
            //include.Xpaths("//*[@id=\"select-input_cover_ncdYears_option_2\"]");
            //include.Sleep();


            ////Have you or any family member included in this policy made any claims within the last 3 years regardless of fault?
            //include.Xpaths("//*[@id=\"bool-input_cover_criminalOffences_false_button\"]");
            //include.Sleep();

            ////Have you or any family member included on this policy had insurance declined, cancelled or any special terms imposed within the past 3 years?
            //include.Xpaths("//*[@id=\"bool-input_cover_insuranceRefused_false_button\"]");
            //include.Sleep();


            ////What is your main telephone number?
            //include.Xpaths2("//*[@id=\"text-input_proposer_telephones\"]", "0761187129");
            //include.Sleep();

            ////Our insurance partners can help you get the right cover at the right price, the providers with the two cheapest quotes may wish to discuss your cover with you. Are you happy to be contacted?
            //include.Xpaths2("//*[@id=\"text-input_email\"]", "achiniiwonder@gmail.com");
            //include.Sleep();

            //include.Xpaths2("//*[@id=\"text-input_password\"]", "AchiniS297");
            //include.Sleep();
            //include.Xpaths("//*[@id=\"question_userRegistration\"]/iw-user-registration/div[1]/iw-question-wrapper/div/div[2]/iw-text-input/div/div[2]/button");
            //include.Sleep();

            ////Renewal reminders
            //include.Xpaths("//*[@id=\"bool-input_proposer_contactPreferences_renewalReminders_false_button\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"bool-input_proposer_contactPreferences_renewalReminders_true_button\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"select-input_carRenewal\"]");
            //include.Sleep();

            //include.Xpaths("//*[@id=\"select-input_carRenewal_option_3\"]");
            //include.Sleep();

            //include.Xpaths("/html/body/app-root/div/div[1]/iw-cycle-cover/div/form/div[2]/div/button");
            //include.Sleep();
        }
    }
}