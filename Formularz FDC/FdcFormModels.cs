using System;
using System.Collections.Generic;
using System.Text;

namespace Formularz_FDC
{
	public class FdcSectionA
	{
		public string date { get; set; }
		public string id { get; set; }
		public string name { get; set; }
		public string surname  { get; set; }
		public FdcSectionA() { }
		public FdcSectionA(string date, string id, string name, string surname)
		{
			this.date = date.ToUpper();
			this.id = id.ToUpper();
			this.name = name.ToUpper();
			this.surname = surname.ToUpper();
		}
	}

	public class FdcSectionB
	{
		public string name { get; set; }
		public string surname  { get; set; }
		public string pesel { get; set; }
		public string birth  { get; set; }
		public string commonGround { get; set; }
		public string idNum  { get; set; }
		public string idType { get; set; }
		public string idId  { get; set; }
		public string relationship  { get; set; }
		public string relationshipId  { get; set; }
		public string disability  { get; set; }
		public string disabilityId { get; set; }

		public string voivodeship  { get; set; }
		public string community  { get; set; }
		public string place  { get; set; }
		public string postal  { get; set; }
		public string street  { get; set; }
		public string flatNum  { get; set; }
		public string apartmentNum  { get; set; }
		public string telephone  { get; set; }
		public FdcSectionB() { }
		public FdcSectionB( string name, string surname, string pesel, string birth, string commonGround, 
							string idNum, string idType, string idId, string relationship, string relationshipId, 
							string disability, string disabilityId, string voivodeship, string community, string place, 
							string postal, string street, string flatNum, string apartmentNum, string telephone )
		{
			this.name = name.ToUpper();
			this.surname = surname.ToUpper();
			this.pesel = pesel.ToUpper();
			this.birth = birth.ToUpper();
			this.commonGround = commonGround.ToUpper();
			this.idNum = idNum.ToUpper();
			this.idType = idType.ToUpper();
			this.idId = idId.ToUpper();
			this.relationship = relationship.ToUpper();
			this.relationshipId = relationshipId.ToUpper();
			this.disability = disability.ToUpper();
			this.disabilityId = disabilityId.ToUpper();
			this.voivodeship = voivodeship.ToUpper();
			this.community = community.ToUpper();
			this.place = place.ToUpper();
			this.postal = postal.ToUpper();
			this.street = street.ToUpper();
			this.flatNum = flatNum.ToUpper();
			this.apartmentNum = apartmentNum.ToUpper();
			this.telephone = telephone.ToUpper();
		}
	}

	public class FdcSectionC
	{
		public string place  { get; set; }
		public string date { get; set; }
		public FdcSectionC() { }
		public FdcSectionC(string date, string place)
		{
			this.place = place.ToUpper();
			this.date = date.ToUpper();
		}
	}
}
