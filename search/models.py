from django.db import models

#案件情報テーブル
class Projectinfo(models.Model):
    jobtype = models.CharField(db_column='jobType', max_length=255, blank=True, null=True)  # Field name made lowercase.
    recruitmentnumbers = models.CharField(db_column='recruitmentNumbers', max_length=255, blank=True, null=True)  # Field name made lowercase.
    period = models.CharField(max_length=255, blank=True, null=True)
    clientname = models.CharField(db_column='clientName', max_length=255, blank=True, null=True)  # Field name made lowercase.
    workcontents = models.CharField(db_column='workContents', max_length=255, blank=True, null=True)  # Field name made lowercase.
    workprocess = models.CharField(db_column='workProcess', max_length=255, blank=True, null=True)  # Field name made lowercase.
    mustskill = models.CharField(db_column='mustSkill', max_length=255, blank=True, null=True)  # Field name made lowercase.
    passableskill = models.CharField(db_column='passableSkill', max_length=255, blank=True, null=True)  # Field name made lowercase.
    skilllevel = models.CharField(db_column='skillLevel', max_length=255, blank=True, null=True)  # Field name made lowercase.
    paymentunitprice = models.CharField(db_column='paymentUnitPrice', max_length=255, blank=True, null=True)  # Field name made lowercase.
    numberofinterviews = models.CharField(db_column='numberOfInterviews', max_length=255, blank=True, null=True)  # Field name made lowercase.
    foreignnationality = models.CharField(db_column='foreignNationality', max_length=255, blank=True, null=True)  # Field name made lowercase.
    bpallowed = models.CharField(db_column='bpAllowed', max_length=255, blank=True, null=True)  # Field name made lowercase.
    startdate = models.CharField(db_column='startDate', max_length=255, blank=True, null=True)  # Field name made lowercase.
    neareststation = models.CharField(db_column='nearestStation', max_length=255, blank=True, null=True)  # Field name made lowercase.
    telework = models.CharField(max_length=255, blank=True, null=True)
    remarks = models.CharField(max_length=255, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'ProjectInfo'

#技術者情報テーブル
class Engineerinfo(models.Model):
    id = models.IntegerField(primary_key=True)
    name = models.CharField(max_length=20)
    age = models.IntegerField(blank=True, null=True)
    gender = models.IntegerField(blank=True, null=True)
    neareststation = models.CharField(db_column='nearestStation', max_length=20, blank=True, null=True)  # Field name made lowercase.
    finaleducation = models.CharField(db_column='finalEducation', max_length=30, blank=True, null=True)  # Field name made lowercase.
    foreignnationality = models.IntegerField(db_column='foreignNationality', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'EngineerInfo'


#スキル情報テーブル
class Skillinfo(models.Model):
    id = models.IntegerField()
    code = models.CharField(max_length=20)
    skillid = models.CharField(db_column='skillId', max_length=6)  # Field name made lowercase.
    skillinfoid = models.CharField(db_column='skillInfoId', primary_key=True, max_length=20)  # Field name made lowercase.
    classification = models.CharField(max_length=20)
    years = models.CharField(max_length=5, blank=True, null=True)
    skilllevel = models.CharField(db_column='skillLevel', max_length=1, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'SkillInfo'


#コードマスタテーブル
class Codemaster(models.Model):
    classification = models.CharField(max_length=2)
    code = models.IntegerField()
    skillid = models.CharField(db_column='skillId', primary_key=True, max_length=6)  # Field name made lowercase.
    name = models.CharField(max_length=50)
    createddate = models.CharField(db_column='createdDate', max_length=20)  # Field name made lowercase.
    updateddate = models.CharField(db_column='updatedDate', max_length=20, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'CodeMaster'