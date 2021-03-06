﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jobs
{
    class Company
    {
        public string Name { get; set; }
        public bool HasPost { get; set; }
        public string Adress { get; set; }
        public string Postal { get; set; }
        public string Phone { get; set; }
        public string CVR { get; set; }
        public string Pnumber { get; set; }
        public string Email { get; set; }
        public string NextAgreementEnd { get; set; }
        public int Students { get; set; }

        public Company()
        {

        }
        public Company(string name, string hasPost, string adress, string postal, string phone, string cvr, string pnumber, string email, string nextAgreementEnd, int Students)
        {
            this.Name = name;
            this.Adress = adress;
            this.Postal = postal;
            this.Phone = phone;
            this.CVR = cvr;
            this.Pnumber = pnumber;
            this.Email = email;
            this.NextAgreementEnd = nextAgreementEnd;
            this.Students = Students;
            //convert string to bool
            if (hasPost == "Nej")
            {
                this.HasPost = false;
            }
            else
            {
                this.HasPost = true;
            }

        }
    }
}
