using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using Castle.Facilities.NHibernateIntegration;
using FluentNHibernate.Query;
using iTextSharp.text;
using iTextSharp.text.factories;
using iTextSharp.text.pdf;
using Mediu.Cms.Domain;
using Mediu.Cms.Domain.Academics;
using Mediu.Cms.Domain.Applications;
using Mediu.Cms.Domain.Applications.Scholarship;
using Mediu.Cms.Domain.CampusEvents;
using Mediu.Cms.Domain.Finances;
using Mediu.Cms.Domain.Infra;
using Mediu.Cms.Domain.PartyRoles;
using Mediu.Cms.Domain.Profiles;
using Mediu.Cms.Domain.RunningNos;
using Mediu.Cms.Domain.StudentRecords;
using Mediu.Cms.Services.Applications;
using Mediu.Cms.Services.Documents;
using Mediu.Cms.Domain.StudentRecords.PostGraduate;
using Mediu.Cms.Services.RunningNos;

namespace Mediu.Cms.Services.Messaging
{
    public class CorrespondenceService : Mediu.Cms.Services.Messaging.ICorrespondenceService
    {
        ISessionManager sm;
        IMessagingService svcMessaging;
        ApplyServiceFactory applyServiceFactory;
        IDocumentServices svcDoc;
        IEmailDocumentServices emailDocumentSvc;
        IRepository<CourseEntryRequirement> courseRequireRepo; //= GlobalApplication.Container.Resolve<IRepository<CourseEntryRequirement>>();
        IThesisReportRunningNoBuilder svcThesisReportRunningNo;
        public CorrespondenceService(ISessionManager sessionManager, IMessagingService messagingService, ApplyServiceFactory applyServiceFactory, IDocumentServices svcDoc, IEmailDocumentServices emailDocumentSvc, IRepository<CourseEntryRequirement> courseRequireRepo, IThesisReportRunningNoBuilder thesisReportRunningNo)
        {
            this.sm = sessionManager;
            this.svcMessaging = messagingService;            
            this.applyServiceFactory = applyServiceFactory;
            this.svcDoc = svcDoc;
            this.emailDocumentSvc = emailDocumentSvc;
            this.courseRequireRepo = courseRequireRepo;
            this.svcThesisReportRunningNo = thesisReportRunningNo;
        }

        public Faculty GetFaculty(string facultyCode)
        {
            return sm.OpenSession().GetOne<Faculty>().Where(x => x.Code).IsEqualTo(facultyCode).Execute();
        }

        public bool OldSendConditionalOfferLetterToApplicant(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "Conditional Offer of Admission";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                            <b>Reference :</b>" + apply.RefNo + @"<br />
                            <b>Name :</b>" + apply.ApplicantName + @" <br />
                            <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                            <p>Based on your academic qualification, we are very pleased to issue a “Conditional Offer of Admission”</p>

                            <p>This offer of admission is conditional upon submitting an additional document required after your 
                                application has been assessed by the Student Admission committee (SAC). <span style=""color: #FF0000; text-decoration: underline; font-weight: bold"">Kindly submit the below 
                                document urgently</span>
                            </p>
                            
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkEn + @"
                               </span> 
                            </p>

                            <p>
                                If you do not submit the required document within 7 days, The University will assume that you are no 
                                longer interested in the programme and University reserves the right to revoke this offer.
                            </p>
                            
                            <p> 
                                For any further clarification please email to <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a> 
                                with your <b><u>Student Reference Number</u></b>. Or you can call our customer service centre at 0060355113939
                            </p>
                            
                            <p>
                                Thank you.
                            </p>
                            
                            <p>
                            Admission & Registration<br>
                            Al-Madinah International University
                            </p>
                                
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>أخي العزيز/ أُختي العزيزة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>
                            نرجو أن يصلك خطابنا هذا في كامل صحتك وإيمانك.
                            استنادا إلى وثائقك الأكاديمية التي قمت بإرسالها إلينا، فإننا سعداء جدا بإبلاغك أن الجامعة قد وافقت على منحك ""خطاب قبول مشروط""  
                            </p>
                            
                            <p>
                            هذا العرض للقبول مشروط بتقديم الوثائق الإضافية المطلوبة بعد أن تم تقييم طلبكم من قبل لجنة القبول. يرجى تقديم الوثائق التالية فيما يلي على وجه السرعة :
                            </p>
                            
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkAr + @"
                               </span> 
                            </p>
                            
                            <p>
                            إذا لم تقم بارسال الوثائق المطلوبة في غضون 7 أيام ، فاننا سنفترض انك لم تعد مهتماً في هذا البرنامج ، ونحن نحتفظ بالحق في إلغاء هذا العرض.
                            </p>
                            
                            <p>
                             ولمزيد من الإستفسار نرجوا التواصل معنا عبر البريد الإلكتروني الآتي: <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a>
                            أو يمكنكم الاتصال المباشر بمركز خدمة العملاء على رقم التلفون 0060355113939 ونرجو تزويدنا برقمك المرجعي  في حال إرسالك لأي استفسار.

                            </p>
                            
                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            قسم القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";

            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public bool SendConditionalOfferLetterToApplicant(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "Conditional Offer Letter";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                            <b>Reference :</b>" + apply.RefNo + @"<br />
                            <b>Name :</b>" + apply.ApplicantName + @" <br /> 
                            
                            <p style='text-align:center;'><h2>Conditional Offer Letter FOR " + apply.Intake.DetailEn + @" INTAKE</h2></p>
                
                            <p style='text-align:center;'>Assalamualaikom warahmatullah wabarakatuh, </p>                           
                            <p>Dear Applicant,</p>
                            <p>Al-Madinah International University (MEDIU) is pleased to inform you that the University has decided to offer you admission to the programme of:</p>
                            
                            <div style='width:99%;border:1px solid black;text-align:center;'><b>" + apply.Course1.NameEn + @"</b></div>
            
                            <p><span style=""text-decoration: underline; "" >provided that you send us the below required documents as soon as possible:</span></p>
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkEn + @"
                               </span> 
                            </p>

                            <p>
                                 We would like to inform you also that if you do not submit the required documents within 14 days, we will assume that you are no longer interested in the programme and University reserves the right to cancel your application.
                            </p>
                            <p>
                                For any further clarification please email to <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a> with your Student Reference Number. Or you can call our Marketing Department in Malaysia at +603 - 5511 3939 Ext : 404 or 319.
                            </p>
                            <p>
                                Thank you. 
                            </p>                            
                            
                            <p>
                            Al-Madinah International University<br/> 
                            Deanship of Admission & Registration    
                            </p>
                                
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p style='text-align:center;'><h2>إشعار قبول المشروط لفصل " + apply.Intake.DetailAr + @"</h2></p>
            
                            <p><i>عزيزي المتقدم؛<br />
                            السلام عليكم ورحمة الله ويركاته
                            </i></p>
                            
                            <p>نشكرك على اهتمامك بالدراسة في جامعة المدينة العالمية، كما يسرنا إشعاركم بأنه قد تم قبولكم قبولاً مشروطًا باستكمال بعض الوثائق للدراسة في برنامج:</p>
                           <br /><br />
                            <div style='width:99%;border:1px solid black;text-align:center;'><b>" + apply.Course1.NameAr + @"</b></div>
            
                            <p>
                             لذا نرجو منك الإسراع بتقديم الوثائق المطلوبة، حتى يتسنى لنا إرسال إشعار القبول لكم في أقرب وقت ممكن، وهي كالتالي:
                            </p>
                            
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkAr + @"
                               </span> 
                            </p>
                            
                             <p>
                                ونود إعلامكم بأنه في حالة عدم قيامكم بإرسال الوثائق المطلوبة في غضون 14 يوما، فسنعدّكم من غير الراغبين بالقبول في هذا البرنامج، وحينئذ يكون لنا الحق في إلغاء طلبكم.
                                ولمزيد من الإستفسار نرجوا التواصل معنا عبر البريد الإلكتروني الآتي: <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a>  أو يمكنكم الاتصال المباشر بقسم التسويق في ماليزيا على رقم التلفون 3939 5511-603+ تحويلة 404 أو 319. ونرجو تزويدنا برقمك المرجعي في حال إرسالك لأي استفسار.

                            </p>
                            
                            <p>
                            شكرا ًجزيلاً،،،،
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            عمادة القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";
            //apply.Applicant.Profile.Email
            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public bool SendDILetterToApplicant(AdmissionApply apply, string remark)
        {
            var subject = "MEDIU Admission";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                            <b>Reference :</b>" + apply.RefNo + @"<br />
                            <b>Name :</b>" + apply.ApplicantName + @" <br />
                            <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                            <p>Thank you for your interest in studying at MEDIU, we are pleased to inform you that after verifying your sent documents we found them uncompleted regarding to the admission requirements at the University.</p>

                            <p>Please submit the required documents as soon as possible, in order to convert your application to the Student Admission Committee (SAC) for final decision on you admission.</p>

                            <p><span style=""color: #FF0000; text-decoration: underline; font-weight: bold"" >The details of required documents as followings:</span></p>
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remark + @"
                               </span> 
                            </p>

                            <p>We would like to inform you also that if you do not submit the required documents within 14 days, we will assume that you are no longer interested in the programme and University reserves the right to cancel your application.</p>
                            
                            <p>For any further clarification please email to admission@mediu.edu.my with your Student Reference Number. Or you can call our Marketing Department in Malaysia at +603 - 5511 3939 Ext : 760 or 761.</p>
                            
                            <p>
                                Thank you.
                            </p>
                            
                            <p>
                            Al-Madinah International University<br/> 
                            Deanship of Admission & Registration    
                            </p>
                                
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>أخي العزيز/ أُختي العزيزة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>نشكرك على اهتمامك بالدراسة في جامعة المدينة العالمية ويسرنا إبلاغك بأننا قد قمنا بتدقيق الوثائق التي أرسلتها فوجدنا أنها غير كاملة حسب متطلبات القبول بالجامعة.</p>
            
                            <p>لذا نرجو منك الإسراع بتقديم الوثائق الإضافية المطلوبة أو إكمال الناقص منها، حتى يتسنى لنا إرسال طلبك إلى لجنة القبول لاتخاذ قرار قبولك.</p>
                            
                            <p>وتفصيل الوثائق المطلوبة كالآتي:</p>
                            
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remark + @"
                               </span> 
                            </p>
                            
                            <p>ونود إعلامك أنه في حالة عدم قيامك بإرسال الوثائق المطلوبة في غضون 14 يوما ، فسنفترض أنك لم تعد مهتماً في هذا البرنامج ، وحينئذ يكون لنا الحق في إلغاء طلبك.</p>
                            
                            <p>ولمزيد من الإستفسار نرجوا التواصل معنا عبر البريد الإلكتروني الآتي:admission@mediu.edu.my أو يمكنكم الاتصال المباشر بقسم التسويق في ماليزيا على رقم التلفون 3939 5511-603+ تحويلة 760 أو 761. ونرجو تزويدنا برقمك المرجعي في حال إرسالك لأي استفسار.</p>                           

                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            عمادة القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";
            //apply.Applicant.Profile.Email
            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public bool SendKIVLetterToApplicant(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "MEDIU Admission";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                            <b>Reference :</b>" + apply.RefNo + @"<br />
                            <b>Name :</b>" + apply.ApplicantName + @" <br />
                            <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                            <p>Thank you for your interest in studying at MEDIU, we are pleased to inform you that after verifying your sent documents we found them uncompleted regarding to the admission requirements at the University.</p>

                            <p>Please submit the required documents as soon as possible, in order to convert your application to the Student Admission Committee (SAC) for final decision on you admission.
                            </p>
                            <p><span style=""color: #FF0000; text-decoration: underline; font-weight: bold"" >The details of required documents as followings:</span></p>
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkEn + @"
                               </span> 
                            </p>

                            <p>
                                 We would like to inform you also that if you do not submit the required documents within 14 days, we will assume that you are no longer interested in the programme and University reserves the right to cancel your application. 
                            </p>
                            
                            <p> 
                                For any further clarification please email to <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a> 
                                with your <b><u>Student Reference Number</u></b>. Or you can call our Marketing Department in Malaysia at +603 - 5511 3939 Ext : 760 or 761.
                            </p>
                            
                            <p>
                                Thank you.
                            </p>
                            
                            <p>
                            Al-Madinah International University<br/> 
                            Deanship of Admission & Registration    
                            </p>
                                
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>أخي العزيز/ أُختي العزيزة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>نشكرك على اهتمامك بالدراسة في جامعة المدينة العالمية ويسرنا إبلاغك بأننا قد قمنا بتدقيق الوثائق التي أرسلتها فوجدنا أنها غير كاملة حسب متطلبات القبول بالجامعة.</p>
            
                            <p>
                            لذا نرجو منك الإسراع بتقديم الوثائق الإضافية المطلوبة أو إكمال الناقص منها، حتى يتسنى لنا إرسال طلبك إلى لجنة القبول لاتخاذ قرار قبولك.
                            </p>
                            <p>
                             وتفصيل الوثائق المطلوبة كالآتي:
                            </p>
                            
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkAr + @"
                               </span> 
                            </p>
                            
                             <p>
                            ونود إعلامك أنه في حالة عدم قيامك بإرسال الوثائق المطلوبة في غضون 14 يوما ، فسنفترض أنك لم تعد مهتماً في هذا البرنامج ، وحينئذ يكون لنا الحق في إلغاء طلبك. 
                            </p>
                            
                            <p>
                             ولمزيد من الإستفسار نرجوا التواصل معنا عبر البريد الإلكتروني الآتي:<a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a>
                            أو يمكنكم الاتصال المباشر بقسم التسويق في ماليزيا على رقم التلفون 3939 5511-603+ تحويلة 760 أو 761.  ونرجو تزويدنا برقمك المرجعي  في حال إرسالك لأي استفسار.

                            </p>                           

                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            عمادة القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";
            //apply.Applicant.Profile.Email
            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public bool SendEmailRequestVisaDocument(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "MEDIU Admission";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                            <b>Reference :</b>" + apply.RefNo + @"<br />
                            <b>Name :</b>" + apply.ApplicantName + @" <br />
                            <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                            <p>Thank you for your interest in studying at MEDIU, as on campus we are pleased to inform you that after verifying your sent documents we found them uncompleted regarding to the immigration requirements.</p>

                            <p>Please submit the required documents as soon as possible, in order to obtain calling visa/ approval from immigration authorities.</p>

                            <p><span style=""color: #FF0000; text-decoration: underline; font-weight: bold"" >The details of required documents as followings:</span></p>
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkEn + @"
                               </span> 
                            </p>

                            <p>
                                 We would like to inform you also that if you do not submit the required documents within 7 days, we will assume that you are no longer interested in the programme and University reserves the right to cancel your application. 
                            </p>
                                                       
                            <p> 
                                For any further clarification please email to <a href=""mailto:visa@mediu.edu.my"">visa@mediu.edu.my</a> with your <b><u>Student Reference Number</u></b>. Or you can directly call customer services in Malaysia at +603 - 5511 3939 and speak to person in charge.
                            </p>
                            <p>
                                Thank you.
                            </p>
                            
                            <p>
                            Al-Madinah International University<br/> 
                            Deanship of Student Affairs 
                            </p>
                                
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>أخي العزيز/ أُختي العزيزة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>نشكرك على اهتمامك بالدراسة في جامعة المدينة العالمية بنظام التعليم المباشر ويسرنا إبلاغكم بأننا قد قمنا بتدقيق الوثائق التي أرسلتموها فوجدنا أنها غير كاملة حسب متطلبات إدارة الهجرة.</p>
                            
                            <p>لذا نرجو منكم الإسراع بتقديم الوثائق الإضافية المطلوبة أو إكمال الناقص منها، حتى يتسنى لنا إرسال طلبك إلى الجهات المختصة لغرض الحصول على الموافقة.</p>

                            <p>
                             وتفصيل الوثائق المطلوبة كالآتي:
                            </p>
                            
                            <p><span style=""color: #FF0000; font-weight: bold"">
                                " + remarkAr + @"
                               </span> 
                            </p>
                            
                            <p>ونود إعلامك أنه في حالة عدم قيامكم بإرسال الوثائق المطلوبة في غضون 7 أيام ، فسنفترض أنك لم تعد مهتماً في هذا البرنامج ، وحينئذ يكون لنا الحق في إلغاء طلبك. </p>
                                            
                            <p>ولمزيد من الإستفسار نرجوا التواصل معنا عبر البريد الإلكتروني الآتي:<a href=""mailto:visa@mediu.edu.my"">visa@mediu.edu.my</a> أو يمكنكم الاتصال المباشر بقسم خدمة العملاء بماليزيا على رقم التلفون 3939 5511-603+ وطلب التحدث إلى الموظف المسئول. ونرجو دائماً تزويدنا برقمك المرجعي في حال إرسالك لأي استفسار. </p>

                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            عمادة شؤون الطلاب  
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";
            //apply.Applicant.Profile.Email
            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public bool SendReplyFormToApplicant(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "Reply Form استمارة القبول";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <b>Reference :</b>" + apply.RefNo + @"<br />
                                <b>Name :</b>" + apply.ApplicantName + @" <br />
                                <p>Dear Applicant, </p>
                                <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                                <p>We congratulate you on the occasion of your admission to study at the Al-Madinah International University, and please use the link below to reply to the admission offer</p>
                                     
                                    <p>http://online.mediu.edu.my/applicantportal/ReplyToOfferLetter?lang=en</p>
                                <p>
                                Thank you
                                </p>
                                <p>
                                Admission & Registration<br>
                                Al-Madinah International University
                                </p>
                                    
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>أخي المتقدم/ أُختي المتقدمة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>نهنئك بمناسبة قبولك للدراسة بجامعة المدينة العالمية، ونرجو منك استخدام الرابط أدناه للرد على عرض القبول</p>

                                <p>http://online.mediu.edu.my/applicantportal/ReplyToOfferLetter?lang=ar</p>

                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            قسم القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";

            //return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }


        public bool SendOfferLetterToApplicant(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "Reply Form استمارة القبول";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <b>Reference :</b>" + apply.RefNo + @"<br />
                                <b>Name :</b>" + apply.ApplicantName + @" <br />
                                <p>Dear Applicant, </p>
                                <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                                <p>We congratulate you on the occasion of your admission to study at the Al-Madinah International University, (MEDIU) the attached file above is your offer letter kindly read its content carefully.</p>
                                <p>After taking your decision please use the link below to reply to the admission offer</p>
                                    
                                    <p>http://online.mediu.edu.my/applicantportal/ReplyToOfferLetter?lang=en</p>
                                <p>
                                Thank you
                                </p>
                                <p>                                
                                Al-Madinah International University
                                <br/>
                                Deanship of Admission & Registration
                                </p>

                                    
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>عزيزي المتقدم/ عزيزتي المتقدمة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>نهنئكم بمناسبة قبولكم للدراسة بجامعة المدينة العالمية، وبرفق هذه الرسالة إشعار قبولكم في الملف المرافق أعلاه، فنرجو الاطلاع عليه وقراءته بتمعن.</p>
                            
                            <p>فبعد اتخاذكم لأي قرار يرجى استخدام الرابط أدناه للرد على عرض القبول.</p>

                                <p>http://online.mediu.edu.my/applicantportal/ReplyToOfferLetter?lang=ar</p>

                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            عمادة القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";

            //return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);

            //Generate Offer Letter and transfer the file over the web service


            var y = new Random().Next(0, 1000);
            string offerLetterFileName = apply.Applicant.UserName + "_OfferLetter" + y.ToString() + ".pdf";
            string offerLetterPath = "";

            if (apply.IsOfferLetterSent)
            {
                //if already generated use the existing file
                //offerLetterPath = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\OfferLetters\\" + apply.Applicant.Id.ToString() + ".pdf";
                //if (File.Exists(offerLetterPath))
                //{
                //    var r = new Random().Next(0, 1000);
                //    string renamedOfferLetterPath = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\OfferLetters\\" + apply.Applicant.Id.ToString() + "_" + r.ToString() + ".pdf";
                //    File.Move(offerLetterPath, renamedOfferLetterPath);
                //    //offerLetterPath = GenerateStudentOfferLetter(apply, offerLetterFileName);
                //}
            }
            offerLetterPath = GenerateStudentOfferLetter(apply, offerLetterFileName, null, null);

            if (File.Exists(offerLetterPath))
            {
                //Send it by Email
                System.IO.BinaryReader br = new System.IO.BinaryReader(System.IO.File.Open(offerLetterPath, System.IO.FileMode.Open, System.IO.FileAccess.Read));
                br.BaseStream.Position = 0;
                byte[] buffer = br.ReadBytes(Convert.ToInt32(br.BaseStream.Length));
                br.Close();
                //apply.Applicant.Profile.Email
                bool result = svcMessaging.SendEmailWithAttachment(apply.Applicant.Profile.Email, subject, body, true, buffer, offerLetterFileName);

                //Take a copy for reviewing after this
                //string copyPath = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\OfferLetters\\";
                string copyPath = @"c:\cms_temp\OfferLetter\";
                //apply.Applicant.Id.ToString() 
                //
                string savefilename = Guid.NewGuid().ToString();
                string copyFile = copyPath + savefilename + ".pdf";
                if (!Directory.Exists(copyPath))
                {
                    Directory.CreateDirectory(copyPath);
                }
                File.Copy(offerLetterPath, copyFile, true);
                apply.LatestOfferLetterFileName = savefilename;

                // Save the pdf file in filestore
                var doc = new Mediu.Cms.Domain.Profiles.Document();
                doc.Data = buffer;
                doc.Category = "STUDENTRECORD_OFFERLETTER";
                doc.Name = offerLetterFileName;
                doc.IsSoftCopyAvailable = true;
                doc.Title = "OFFER LETTER";
                doc.Description = doc.Title + " for " + apply.Applicant.AdmissionApply.RefNo;
                doc.MimeType = "application/pdf";
                svcDoc.SaveDocument(apply.Applicant.Profile, doc);

                return result;
            }
            else
            {
                return false;
            }
        }

        public bool SendVisaApprovalToApplicant(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "Reply Form استمارة القبول";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <b>Reference :</b>" + apply.RefNo + @"<br />
                                <b>Name :</b>" + apply.ApplicantName + @" <br />
                                <p>Dear Applicant, </p>
                                <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                                <p>We congratulate you on the occasion of your admission to study at the Al-Madinah International University, (MEDIU) the attached file above is your offer letter kindly read its content carefully.</p>
                                <p>After taking your decision please use the link below to reply to the admission offer</p>
                                    
                                    <p>http://online.mediu.edu.my/applicantportal/ReplyToOfferLetter?lang=en</p>
                                <p>
                                Thank you
                                </p>
                                <p>                                
                                Al-Madinah International University
                                <br/>
                                Deanship of Admission & Registration
                                </p>

                                    
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>عزيزي المتقدم/ عزيزتي المتقدمة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>نهنئكم بمناسبة قبولكم للدراسة بجامعة المدينة العالمية، وبرفق هذه الرسالة إشعار قبولكم في الملف المرافق أعلاه، فنرجو الاطلاع عليه وقراءته بتمعن.</p>
                            
                            <p>فبعد اتخاذكم لأي قرار يرجى استخدام الرابط أدناه للرد على عرض القبول.</p>

                                <p>http://online.mediu.edu.my/applicantportal/ReplyToOfferLetter?lang=ar</p>

                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            عمادة القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";

            //return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);

            //Generate Offer Letter and transfer the file over the web service

            var docList = apply.Applicant.Profile.Documents.Where(x => x.Category == "Visa Approval").LastOrDefault();


            if (docList.Data != null)
            {
                //Send it by Email

                //apply.Applicant.Profile.Email
                bool result = svcMessaging.SendEmailWithAttachment(apply.Applicant.Profile.Email, subject, body, true, docList.Data, docList.Name);

                return result;
            }
            else
            {
                return false;
            }
        }

        public bool PreviewOfferLetter(AdmissionApply apply, string remarkEn, string remarkAr, DateTime? dateOfferLetterSent)
        {
            //Generate Offer Letter 


            var y = new Random().Next(0, 1000);
            string offerLetterFileName = apply.Applicant.UserName + "_OfferLetter" + y.ToString() + ".pdf";
            string offerLetterPath = "";

            offerLetterPath = GenerateStudentOfferLetter(apply, offerLetterFileName, dateOfferLetterSent, remarkEn);

            if (File.Exists(offerLetterPath))
            {
                //Send it by Email
                System.IO.BinaryReader br = new System.IO.BinaryReader(System.IO.File.Open(offerLetterPath, System.IO.FileMode.Open, System.IO.FileAccess.Read));
                br.BaseStream.Position = 0;
                byte[] buffer = br.ReadBytes(Convert.ToInt32(br.BaseStream.Length));
                br.Close();

                //Take a copy for reviewing after this
                //string copyPath = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\OfferLetters\\";
                string copyPath = @"c:\cms_temp\OfferLetter\";
                string savefilename = Guid.NewGuid().ToString();
                string copyFile = copyPath + savefilename + ".pdf";
                if (!Directory.Exists(copyPath))
                {
                    Directory.CreateDirectory(copyPath);
                }
                File.Copy(offerLetterPath, copyFile, true);

                apply.LatestOfferLetterFileName = savefilename;

                if (remarkEn == "RegenerateOfferLetter")
                {

                    // Save the pdf file in filestore
                    var doc = new Mediu.Cms.Domain.Profiles.Document();
                    doc.Data = buffer;
                    doc.Category = "STUDENTRECORD_OFFERLETTER";
                    doc.Name = offerLetterFileName;
                    doc.IsSoftCopyAvailable = true;
                    doc.Title = "OFFER LETTER";
                    doc.Description = "Regenerated " + doc.Title + " : " + apply.Applicant.AdmissionApply.RefNo;
                    doc.MimeType = "application/pdf";
                    svcDoc.SaveDocument(apply.Applicant.Profile, doc);

                }

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Faris : Special For Postgraduate For Particular Purpose
        /// </summary>
        /// <param name="post">Postgrad Object</param>
        /// <param name="type">type of letter</param>
        /// <returns></returns>
        public bool PreviewPostgraduateOfficialLetter(PostGraduateThesis post, string type, bool isRinggit)
        {
            var student = post.Student;
            var profile = student.Profile;
            var y = new Random().Next(0, 1000);
            string offerLetterFileName = student.UserName + "_Official" + type + "Letter" + y.ToString() + ".pdf";
            string offerLetterPath = "";

            //we get the actual file that save at temp folder
            offerLetterPath = GeneratePostgradOfficialLetter(post, offerLetterFileName, type.ToLower(), isRinggit);

            if (File.Exists(offerLetterPath))
            {
                //Faris : get file again and convert to byte[]
                using (var br = new System.IO.BinaryReader(System.IO.File.Open(offerLetterPath, System.IO.FileMode.Open, System.IO.FileAccess.Read)))
                {
                    br.BaseStream.Position = 0;
                    byte[] buffer = br.ReadBytes(Convert.ToInt32(br.BaseStream.Length));
                    br.Close();
                    string category = "";
                    string desc = "";
                    switch (type.ToLower())
                    {
                        case "supervisor":
                            category = "OFFICIAL_SUPERVISOR_LETTER";
                            desc = "Official Supervisor Appointed Letter";
                            break;
                        case "chairman":
                            category = "OFFICIAL_CHAIRMAN_LETTER";
                            desc = "Official Chairman Appointed Letter";
                            break;
                        case "internalexaminer1":
                            category = "OFFICIAL_INTERNALEXAMINER1_LETTER";
                            desc = "Official InternalExaminer 1 Appointed Letter";
                            break;
                        case "internalexaminer2":
                            category = "OFFICIAL_INTERNALEXAMINER2_LETTER";
                            desc = "Official InternalExaminer 2 Appointed Letter";
                            break;
                        case "externalexaminer1":
                            category = "OFFICIAL_EXTERNALEXAMINER1_LETTER";
                            desc = "Official ExternalExaminer 1 Appointed Letter";
                            break;
                        case "externalexaminer2":
                            category = "OFFICIAL_EXTERNALEXAMINER2_LETTER";
                            desc = "Official ExternalExaminer 2 Appointed Letter";
                            break;
                        case "depsfacultyrep":
                            category = "OFFICIAL_DEPFACREP_LETTER";
                            desc = "Official DepFacRep Appointed Letter";
                            break;
                        case "studentnotice":
                            category = "OFFICIAL_STUDENTNOTICE_LETTER";
                            desc = "Official Student Appointed Supervisor Letter";
                            break;
                    }
                    // Save the pdf file in filestore
                    var doc = new Mediu.Cms.Domain.Profiles.Document();
                    doc.Data = buffer;
                    doc.Category = category;
                    doc.Name = offerLetterFileName;
                    doc.IsSoftCopyAvailable = true;
                    doc.Title = "Offical Letter";
                    doc.Description = desc;
                    doc.MimeType = "application/pdf";
                    svcDoc.SaveDocument(profile, doc);
                    File.Delete(offerLetterPath);
                    return true;
                }
            }
            else
            {
                return false;
            }

        }

        public void GenerateUniversialOfferLetter(string destOffer, AdmissionApply apply, string offerLetterFileName, DateTime? dateOfferLetterSent)
        {
            DateTime currentDateTime = (dateOfferLetterSent != null) ? (DateTime)dateOfferLetterSent : DateTime.Now;
            var s = sm.OpenSession();

            AdmissionApply absApply = apply;
            var country = s.GetAll<Country>().Where(c => c.Code).IsEqualTo(absApply.Applicant.Profile.Citizenship).Execute().FirstOrDefault();
            string studyMode = absApply.LearningMode.ToString();
            string sourceOffer = "";
            //control Offer Letter Templete
            if (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONCAMPUS")
            {
                if (absApply.Course1.CourseDescription.CourseLevel == "B" || absApply.Course1.CourseDescription.CourseLevel == "D")
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Undergraduate_auto.pdf";
                else
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Postgraduate_auto.pdf";
            }
            else
            {
                if (absApply.Course1.CourseDescription.CourseLevel == "B" || absApply.Course1.CourseDescription.CourseLevel == "D")
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Undergraduate_auto_MY.pdf";
                else
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Postgraduate_auto_MY.pdf";
            }

            try
            {
                IntakeEvent semester = apply.Intake;
                var semesters = s.GetAll<CampusEvent>()
                                            .ThatHasChild(c => c.ParentEvent)
                                                .Where(c => c.Id).IsEqualTo(semester.ParentEvent.Id)
                                            .EndChild()
                                            .And(c => c.EventType).IsEqualTo("SME")
                                            .Execute().ToList();
                CampusEvent semesterDates = semesters.Where(c => c.Code.ToLower().StartsWith(semester.Code.Substring(0, 3).ToLower())).FirstOrDefault();

                int serial = semester.OfferLetterSerialNumber + 1;
                string cardnum = absApply.Applicant.Profile.IdCardNumber;
                if (cardnum == "nil")
                {
                    cardnum = "";
                }

                PdfReader r = new PdfReader(sourceOffer);
                string guid = Guid.NewGuid().ToString();
                PdfStamper stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate));
                PdfContentByte canvas = stamper.GetOverContent(1);
                PdfContentByte canvas2 = stamper.GetOverContent(2);
                PdfContentByte canvas3 = stamper.GetOverContent(3);
                BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                canvas.SetFontAndSize(bf, 8);
                canvas2.SetFontAndSize(bf, 8);

                Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);

                //Calculate admin fee
                decimal totalMyr = 0;
                var immigrationBondFee = s.GetAll<ImmigrationBondFeeGroup>().Where(c => c.ImmigrationCountryGroup).IsEqualTo(country.ImmigrationCountryGroup).Execute().FirstOrDefault();
                List<FeeStructureItem> applicantFeeStructureItems = new List<FeeStructureItem>();
                var courseGroupItem = s.GetOne<CourseGroupItem>().Where(x => x.Course).IsEqualTo(apply.Course1)
                    .AndHasChild(c => c.CourseGroup).Where(x => x.IsActive).IsEqualTo(true).EndChild().Execute();
                if (courseGroupItem != null)
                {
                    var feeStructure = s.GetOne<FeeStructure>().Where(x => x.CourseGroup).IsEqualTo(courseGroupItem.CourseGroup).Execute();
                    foreach (var item in feeStructure.FeeStructureItems.Where(x => x.Stage == Stage.One))
                    {
                        if (apply.LearningMode == LearningMode.OnCampus)
                        {
                            if (apply.Applicant.Profile.Citizenship == "MY" || studyMode.ToUpper() == "ONLINE")
                            {
                                if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.Local || item.StudentType == StudentType.Both))
                                {
                                    applicantFeeStructureItems.Add(item);
                                }
                            }
                            else
                            {
                                if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.International || item.StudentType == StudentType.Both))
                                {
                                    applicantFeeStructureItems.Add(item);
                                }
                            }
                        }
                        else
                        {
                            if (item.Mode == Mode.Online || item.Mode == Mode.Both)
                            {
                                applicantFeeStructureItems.Add(item);
                            }
                        }
                    }

                    if (applicantFeeStructureItems != null)
                        foreach (var fs in applicantFeeStructureItems.OrderBy(c => c.Name))
                        {

                            totalMyr = totalMyr + Convert.ToDecimal(fs.AmountMYR);

                        }
                }

                //OL/OC
                //PT/WPT apply.PlacementTestExamRequired
                //MAL/INT absApply.Applicant.Profile.Citizenship
                //costlevel apply.Course1.CourseDescription.CourseLevel

                //--- control fees

                //on campus
                decimal pay = 2000;
                string visaFee = "370";
                string medical = "350";
                string unibond = "1000";
                string immiBond = immigrationBondFee.AmountMYR.ToString("0");
                string totalFee = "Five thousand (RM 5000) Malaysian Ringgit, this will include:";
                string totalFeeAr = "خمسة آلاف (5000) رنجيت ماليزي";


                decimal balance = pay - totalMyr;

                if (absApply.PlacementTestExamRequired)
                {
                    balance = balance - 100m;

                }


                string balanceEN = "* Balance RM " + balance.ToString("00") + " will calculated in the student’s payment based on the invoices issued.";
                string balanceAr = "* " + " المتبقي من المبلغ المقدم" + balance.ToString("00") + "يتم احتسابه ضمن مدفوعات الطالب حسب ما يصدر له من فواتير";

                if (absApply.Applicant.Profile.Citizenship == "MY")
                {
                    totalFee = "Three thousand (RM 3000) Malaysian Ringgit";
                    totalFeeAr = "ثلاثة آالاف (3000) رنجيت ماليزي";

                    if (absApply.PlacementTestExamRequired)
                    {
                        totalFee = "Four hundred forty five (RM 445) Malaysian Ringgit";
                        totalFeeAr = "أربعمائة وخمس وأربعون (445) رنجيت ماليزي";
                    }

                    if (absApply.Course1.CourseDescription.CourseLevel == "M")
                    {
                        totalFee = "Three thousand (RM 3000) Malaysian Ringgit";
                        totalFeeAr = "ثلاثة آالاف (3000) رنجيت ماليزي";
                        // totalFee = "Three hundred seventy (RM370) Malaysian Ringgit";
                        // totalFeeAr = "ثلاثمئة وسبعون (370) رنجيت ماليزي";

                        if (absApply.PlacementTestExamRequired)
                        {
                            totalFee = "Four hundred seventy (RM470) Malaysian Ringgit";
                            totalFeeAr = "أربعمائة وسبعون (470) رنجيت ماليزي";
                        }

                    }
                    if (absApply.Course1.CourseDescription.CourseLevel == "P")
                    {
                        //  totalFee = "Three hundred twenty (RM320) Malaysian Ringgit";
                        //  totalFeeAr = "ثلاثمئة مائة وعشرين (320) رنجيت ماليزي";

                        totalFee = "Three thousand (RM 3000) Malaysian Ringgit";
                        totalFeeAr = "ثلاثة آالاف (3000) رنجيت ماليزي";

                        if (absApply.PlacementTestExamRequired)
                        {
                            totalFee = "Four hundred twenty (RM420) Malaysian Ringgit";
                            totalFeeAr = "أربعمائة مائة وعشرين (420) رنجيت ماليزي";
                        }

                    }
                    visaFee = "";
                    medical = "";
                    unibond = "";
                    immiBond = "";
                }

                //online 
                if (studyMode.ToUpper() == "ONLINE")
                {
                    visaFee = "";
                    medical = "";
                    unibond = "";
                    immiBond = "";

                    totalFee = "Three thousands (RM 3000) Malaysian Ringgit, this will include:";
                    totalFeeAr = " ثلاثة آلاف (3000) رنجيت ماليزي ";


                    //if (absApply.Applicant.Profile.Citizenship != "MY")
                    //{

                    //}
                    //balance = "1705";

                    //if (absApply.Applicant.Profile.Citizenship == "MY" && absApply.Course1.CourseDescription.CourseLevel == "B")
                    //{

                    //    totalFee = "Three hundred ninety five(RM 395) Malaysian Ringgit";
                    //    totalFeeAr = "ثلاثمئة وخمسة وتسعون (395) رنجيت ماليزي";

                    //}

                }

                string ptRequired = "MAPT/MEPT  (additional RM 100  is required as test fee)";
                string ptRequiredAr = "امتحان الكفاءة في اللغة العربية/ الإنجليزية (رسوم الامتحان 100 رنجيت ماليزي)";
                string ptRequiredFee = "100";

                //Placement Test Exam Required
                if (!absApply.PlacementTestExamRequired)
                {
                    ptRequired = "N/A";
                    ptRequiredAr = "N/A";
                    ptRequiredFee = "N/A";

                    if (studyMode.ToUpper() == "ONLINE" && absApply.Applicant.Profile.Citizenship == "MY" && absApply.Course1.CourseDescription.CourseLevel == "P" && absApply.Course1.CourseDescription.CourseLevel == "M")
                    {

                        totalFee = "Three thousand (RM 3000) Malaysian Ringgit.";
                        totalFeeAr = " ثلاثة آلاف (3000) رنجيت ماليزي ";

                    }
                }
                if (absApply.Course1.CourseDescription.CourseLevel == "D")
                {
                    totalFee = "Malaysian Ringgit : " + totalMyr.ToString("00");
                    totalFeeAr = "رنجيت ماليزي" + totalMyr.ToString("00");
                }


                //----

                //English
                //A Learning Mode OL / OC

                //ColumnText.ShowTextAligned(canvas,
                //      Element.ALIGN_CENTER, new Phrase(studyMode.ToUpper(), f5), 402, 727, 0);

                // 1 Semester
                string semesterName = "ADMISSION FOR " + semester.MonthName.ToUpper() + " " + semester.Year.ToString() + " ( " + studyMode.ToUpper() + " )";

                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_CENTER, new Phrase(semesterName, f5), 286, 727, 0);

                // 2 Serial Number
                string serialNumber = serial.ToString("0000000") + "(" + currentDateTime.Year.ToString().Substring(2, 2) + currentDateTime.Month.ToString("00") + ")";
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);

                // 3 Reference Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.RefNo, f3), 190, 694, 0);

                // 4 Name
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.NameEnglish.ToUpper(), f3), 190, 682, 0);

                // 5 Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(currentDateTime.ToString("dd.MM.yyyy"), f3), 190, 670, 0);

                // 6 Passport Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.IdCardNumber, f3), 190, 660, 0);

                // 7 Program
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Course1.NameEn, f3), 190, 648, 0);

                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f3), 190, 637, 0);

                //PT 
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(Convert.ToString(ptRequired), f3), 190, 626, 0);

                // 9 Vertual Account Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.HSBCAccPayor, f3), 190, 615, 0);
                //ColumnText.ShowTextAligned(canvas,
                //      Element.ALIGN_LEFT, new Phrase("N/A", f3), 190, 615, 0);

                //Admin fees
                ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_LEFT, new Phrase(totalMyr.ToString("00"), f4e), 228, 483, 0);

                if (absApply.Applicant.Profile.Citizenship == "MY" || (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONLINE"))
                {
                    //Placement fees
                    ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(ptRequiredFee, f4e), 228, 471, 0);
                }

                if (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONCAMPUS")
                {
                    //Visa fees
                    ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(visaFee, f4e), 228, 471, 0);

                    //Immigration Bond Fee
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(immiBond, f4e), 228, 460, 0);

                    //Medical Insurance
                    ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(medical, f4e), 228, 449, 0);

                    //University Bond
                    ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(unibond, f4e), 228, 437, 0);

                }
                // 10 Total Fee
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(totalFee, f4e), 184, 518, 0);

                if (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONLINE")
                {
                    //Balance
                    ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_CENTER, new Phrase(balanceEN, f4e), 300, 461, 0);
                }

                //Arabic -----------------------------------------------------------------------------------------------------------------------               


                // 1 Serial Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f2), 450, 800, 0);


                //A Learning Mode OL / OC
                //oncampus - التعليم المباشر
                // online - التعليم عن بعد
                string studyModeAr = "( التعليم عن بعد )";
                if (studyMode.ToUpper() == "ONCAMPUS")
                {
                    studyModeAr = "( التعليم المباشر )";
                }
                else
                {
                    studyModeAr = "( التعليم عن بعد )";
                }

                //ColumnText.ShowTextAligned(canvas2,
                //        Element.ALIGN_CENTER, new Phrase(studyModeAr, f4), 240, 728, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                String arabicSemesterName = " إشعار قبول مبدئي لفصل " + semester.MonthNameAr + " " + semester.Year.ToString() + " " + studyModeAr;

                // 5 Semester
                float seTextLength = canvas2.GetEffectiveStringWidth(arabicSemesterName, true);

                float position = ((60 - seTextLength) / 2) + 270;
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_CENTER, new Phrase(arabicSemesterName, f4), 299, 728, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                // 2 Reference Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.RefNo, f2), 413, 695, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 3 Name
                string ArabicName = absApply.Applicant.Profile.NameAr;
                if (String.IsNullOrEmpty(ArabicName) || ArabicName == "None")
                {
                    ArabicName = absApply.Applicant.Profile.Name;
                }

                float textLength = canvas2.GetEffectiveStringWidth(ArabicName, true);

                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(ArabicName, f2), 413, 681, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4 Date
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(currentDateTime.ToString("dd MMMMMM yyyy", new CultureInfo("ar-QA")), f2), 413, 667, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4.5 Passport Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(cardnum, f2), 413, 653, 0);



                // 6 Program
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.Course1.NameAr, f2), 413, 639, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f2), 413, 625, 0);


                // 8.2 pt Required Ar
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(ptRequiredAr, f2), 413, 611, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                // 9 Vertual Account Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.HSBCAccPayor, f2), 413, 597, 0);
                //ColumnText.ShowTextAligned(canvas2,
                //     Element.ALIGN_RIGHT, new Phrase("N/A", f2), 413, 597, 0);



                //Admin fees
                ColumnText.ShowTextAligned(canvas2,
                        Element.ALIGN_LEFT, new Phrase(totalMyr.ToString("00"), f4e), 336, 479, 0);

                if (absApply.Applicant.Profile.Citizenship == "MY" || (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONLINE"))
                {
                    //Placement test fee
                    ColumnText.ShowTextAligned(canvas2,
                             Element.ALIGN_LEFT, new Phrase(ptRequiredFee, f4e), 336, 464, 0);
                }
                if (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONCAMPUS")
                {
                    //Visa fees
                    ColumnText.ShowTextAligned(canvas2,
                            Element.ALIGN_LEFT, new Phrase(visaFee, f4e), 336, 464, 0);

                    // 9 Immigration Bond Fee
                    ColumnText.ShowTextAligned(canvas2,
                          Element.ALIGN_LEFT, new Phrase(immiBond, f4e), 336, 451, 0);


                    //Medical Insurance
                    ColumnText.ShowTextAligned(canvas2,
                            Element.ALIGN_LEFT, new Phrase(medical, f4e), 336, 438, 0);

                    //University Bond
                    ColumnText.ShowTextAligned(canvas2,
                            Element.ALIGN_LEFT, new Phrase(unibond, f4e), 336, 423, 0);

                    //Third Page Serial Number
                    ColumnText.ShowTextAligned(canvas3,
                         Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);
                }


                // 10 Total Fee
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(totalFeeAr, f4a), 180, 525, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                if (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONLINE")
                {
                    //Balance
                    ColumnText.ShowTextAligned(canvas2,
                            Element.ALIGN_CENTER, new Phrase(balanceAr, f4e), 300, 453, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                }

                stamper.Close();

            }
            catch (Exception ex)
            {

            }
        }

        public void GenerateOfficialOfferLetter(string destOffer, AdmissionApply apply, string offerLetterFileName, DateTime? dateOfferLetterSent, string blankPrint)
        {
            DateTime currentDateTime = (dateOfferLetterSent != null) ? (DateTime)dateOfferLetterSent : DateTime.Now;
            var s = sm.OpenSession();

            AdmissionApply absApply = apply;
            var country = s.GetAll<Country>().Where(c => c.Code).IsEqualTo(absApply.Applicant.Profile.Citizenship).Execute().FirstOrDefault();
            string studyMode = absApply.LearningMode.ToString();
            string sourceOffer = "";
            Image image;
            Image imageAr;


            //Control Offer Letter Templete
            if (blankPrint != null && blankPrint == "PrintBlank")
            {

                if (studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY")
                {
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Blank_OC.pdf";
                }
                else
                {
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Blank_OL.pdf";
                }

            }
            else
            {
                if (studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY")
                {
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_OC.pdf";
                }
                else
                {
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_OL.pdf";
                }
            }


            //Control signature
            if (absApply.Course1.CourseDescription.CourseLevel == "M" || absApply.Course1.CourseDescription.CourseLevel == "P")
            {
                if (blankPrint != null && blankPrint == "PrintBlank")
                {
                    image = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\doukoureSignBlank.jpg");
                    image.SetAbsolutePosition(59, 255);
                    image.ScaleToFit(150, 110);

                    imageAr = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\doukoureSignArBlank.jpg");
                    imageAr.SetAbsolutePosition(380, 250);
                    imageAr.ScaleToFit(150, 110);
                }
                else
                {
                    image = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\doukoureSign.jpg");
                    image.SetAbsolutePosition(59, 255);
                    image.ScaleToFit(150, 110);

                    imageAr = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\doukoureSignAr.jpg");
                    imageAr.SetAbsolutePosition(380, 250);
                    imageAr.ScaleToFit(150, 110);
                }
            }
            else
            {

                if (blankPrint != null && blankPrint == "PrintBlank")
                {

                    image = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\mohammedSignBlank.jpg");
                    image.SetAbsolutePosition(59, 250);
                    image.ScaleToFit(180, 170);

                    imageAr = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\mohammedSignArBlank.jpg");
                    imageAr.SetAbsolutePosition(380, 250);
                    imageAr.ScaleToFit(150, 110);
                }
                else
                {
                    image = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\mohammedSign.jpg");
                    image.SetAbsolutePosition(59, 250);
                    image.ScaleToFit(180, 170);

                    imageAr = Image.GetInstance(HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Images\\mohammedSignAr.jpg");
                    imageAr.SetAbsolutePosition(380, 270);
                    imageAr.ScaleToFit(150, 110);
                }

            }

            try
            {

                IntakeEvent semester = apply.Intake;
                var semesters = s.GetAll<CampusEvent>()
                                            .ThatHasChild(c => c.ParentEvent)
                                                .Where(c => c.Id).IsEqualTo(semester.ParentEvent.Id)
                                            .EndChild()
                                            .And(c => c.EventType).IsEqualTo("SME")
                                            .Execute().ToList();
                CampusEvent semesterDates = semesters.Where(c => c.Code.ToLower().StartsWith(semester.Code.Substring(0, 3).ToLower())).FirstOrDefault();

                string yearAr = "";
                string monthAr = "";
                string semesterIntake = semester.MonthName.ToUpper() + " " + semester.Year.ToString();

                if (absApply.Course1.CourseDescription.CourseLevel == "P" && absApply.StartMonth != null)
                {

                    semesterIntake = absApply.StartMonth.ToUpper();

                    yearAr = apply.StartMonth.Substring(apply.StartMonth.Length - 4);
                    monthAr = apply.StartMonth.Substring(0, apply.StartMonth.Length - 5);

                    if (monthAr == "January")
                        monthAr = "يناير";
                    else if (monthAr == "February")
                        monthAr = "فبراير";
                    else if (monthAr == "March")
                        monthAr = "مسيرة";
                    else if (monthAr == "April")
                        monthAr = "أبريل";
                    else if (monthAr == "May")
                        monthAr = "قد";
                    else if (monthAr == "June")
                        monthAr = "يونيو";
                    else if (monthAr == "July")
                        monthAr = "يوليو";
                    else if (monthAr == "August")
                        monthAr = "أغسطس";
                    else if (monthAr == "September")
                        monthAr = "سبتمبر";
                    else if (monthAr == "October")
                        monthAr = "أكتوبر";
                    else if (monthAr == "November")
                        monthAr = "نوفمبر";
                    else if (monthAr == "December")
                        monthAr = "ديسمبر";
                    else
                        monthAr = "";
                }


                int serial = semester.OfferLetterSerialNumber + 1;
                string cardnum = absApply.Applicant.Profile.IdCardNumber;
                if (cardnum == "nil")
                {
                    cardnum = "";
                }

                PdfReader r = new PdfReader(sourceOffer);
                string guid = Guid.NewGuid().ToString();
                PdfStamper stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate));
                PdfContentByte canvas = stamper.GetOverContent(1);
                PdfContentByte canvas2 = stamper.GetOverContent(2);
                PdfContentByte canvas3 = stamper.GetOverContent(3);
                PdfContentByte canvas4 = stamper.GetOverContent(4);
                BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                canvas.SetFontAndSize(bf, 8);
                canvas2.SetFontAndSize(bf, 8);
                if (studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY")
                {
                    canvas3.SetFontAndSize(bf, 8);

                }
                Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);

                //Calculate admin fee
                var immigrationBondFee = s.GetAll<ImmigrationBondFeeGroup>().Where(c => c.ImmigrationCountryGroup).IsEqualTo(country.ImmigrationCountryGroup).Execute().FirstOrDefault();
                List<FeeStructureItem> applicantFeeStructureItems = new List<FeeStructureItem>();
                var courseGroupItem = s.GetOne<CourseGroupItem>().Where(x => x.Course).IsEqualTo(apply.Course1)
                    .AndHasChild(c => c.CourseGroup).Where(x => x.IsActive).IsEqualTo(true).EndChild().Execute();

                //checking apply course - 21/8/2017 : using new fee structure
              
                //var courseApplyFee = s.GetOne<FeeStructure>().Where(x => x.Course).IsEqualTo(apply.Course1)     
                //   .AndHasChild(c => c.Course).Where(x => x.IsActive).IsEqualTo(true)
                //   .EndChild().Execute();

               var courseApplyFee = s.GetAll<FeeStructure>().Where(c => c.FeeVersion).IsNotNull().And(c => c.Course).IsEqualTo(apply.Course1).AndHasChild(c=>c.Course).Where(a=>a.IsActive).IsEqualTo(true)
                   .EndChild()
                   .Execute().LastOrDefault();


                if (studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY")
                {
                    if(courseApplyFee != null)
                    //if (courseGroupItem != null)
                    {
                        //var feeStructure = s.GetOne<FeeStructure>().Where(x => x.CourseGroup).IsEqualTo(courseGroupItem.CourseGroup).Execute();
                        var feeStructure = s.GetAll<FeeStructure>().Where(x => x.Course).IsEqualTo(courseApplyFee.Course).Execute().LastOrDefault();
                        foreach (var item in feeStructure.FeeStructureItems.Where(x => x.Stage == Stage.One))
                        {
                            if (apply.LearningMode == LearningMode.OnCampus)
                            {
                                if (apply.Applicant.Profile.Citizenship == "MY" || studyMode.ToUpper() == "ONLINE")
                                {
                                    if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.Local || item.StudentType == StudentType.Both))
                                    {
                                        applicantFeeStructureItems.Add(item);
                                    }
                                }
                                else
                                {
                                    if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.International || item.StudentType == StudentType.Both))
                                    {

                                        //Added to dinamicly get the Amount for ImmigrationBondGroup
                                        if (item.Mode == Mode.OnCampus && item.StudentType == StudentType.International && item.RevenueCategory == RevenueCategory.ImmigrationBond && item.Stage == Stage.One)
                                        {
                                            item.AmountMYR = immigrationBondFee.AmountMYR;
                                            item.AmountUSD = immigrationBondFee.AmountUSD;
                                            applicantFeeStructureItems.Add(item);
                                        }
                                        else
                                        {
                                            applicantFeeStructureItems.Add(item);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (item.Mode == Mode.Online || item.Mode == Mode.Both)
                                {
                                    applicantFeeStructureItems.Add(item);
                                }
                            }
                        }
                    }
                }


                //--- Set additional fees

                string immiBond = immigrationBondFee.AmountMYR.ToString("0");
                List<FeeStructureItem> othersFees = new List<FeeStructureItem>()
                                                         {new FeeStructureItem() { Id=new Guid(), Name = "MAPT/MEPT Fee", NameAr = "امتحان الكفاءة في اللغة العربية/ الإنجليزية", AmountMYR = 100 },
                                                          //new FeeStructureItem() { Id=new Guid(), Name = "Calling Visa Fee", NameAr = "رسوم الفيزا" , AmountMYR = 370}, 
                                                          //new FeeStructureItem() { Id=new Guid(), Name = "Insurance Fee", NameAr = "رسوم التأمين الصحي" , AmountMYR = 350}, 
                                                          new FeeStructureItem() { Id=new Guid(), Name = "University Bond Fee", NameAr = "ضمان الجامعة" , AmountMYR = 1000} 
                                                         //new FeeStructureItem() { Id=new Guid(), Name = "Immigration Bond Fee", NameAr = "رسوم تأمين الجوازات" , AmountMYR = immigrationBondFee.AmountMYR}
                                                          
                                                          };






                //English


                // 1 Semester
                string semesterName = "ADMISSION FOR " + semesterIntake + " ( " + studyMode.ToUpper() + " )";

                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_CENTER, new Phrase(semesterName, f5), 286, 727, 0);

                // 2 Serial Number
                string serialNumber = serial.ToString("0000000") + "(" + currentDateTime.Year.ToString().Substring(2, 2) + currentDateTime.Month.ToString("00") + ")";
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);

                // 3 Reference Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.RefNo, f3), 190, 694, 0);

                // 4 Name
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.NameEnglish.ToUpper(), f3), 190, 682, 0);

                // 5 Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(currentDateTime.ToString("dd.MM.yyyy"), f3), 190, 670, 0);

                // 6 Passport Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.IdCardNumber, f3), 190, 660, 0);

                // 6.A Nationality
                var nationality = s.GetAll<Country>().Where(c => c.Code).IsEqualTo(absApply.Applicant.Profile.Citizenship).Execute().FirstOrDefault();

                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(nationality.NameEn.ToUpper(), f3), 190, 648, 0);

                // 7 Program
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Course1.NameEn, f3), 190, 637, 0);

                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f3), 190, 627, 0);

                var dur = absApply.Course1.CourseDescription.DurationType;
                string durationType = "";
                string durationTypeAr = "";
                if (dur == DurationType.Months)
                {
                    durationType = "Months";
                    durationTypeAr = "(أشهر)";
                }
                else if (dur == DurationType.Weeks)
                {
                    durationType = "Weeks";
                    durationTypeAr = "(أسابيع)";
                }
                else
                {
                    durationType = "Years";
                    durationTypeAr = "(سنة/سنوات)";
                }
                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(durationType, f3), 200, 627, 0);

                // 9 Vertual Account Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.HSBCAccPayor, f3), 190, 615, 0);//open on 14 August 2017

                //Extra Requirement List 
                int erListIndexEn = 604;
                int erListIndexAr = 585;


                if (absApply.PlacementTestExamRequired)
                {
                    string ptRequired = "MAPT/MEPT  (additional RM 100  is required as test fee).";
                    string ptRequiredAr = "امتحان الكفاءة في اللغة العربية/ الإنجليزية (رسوم الامتحان 100 رنجيت ماليزي).";

                    //Placement Test Exam Required
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(Convert.ToString(ptRequired), f3), 190, erListIndexEn, 0);

                    //  Placement Test Exam Required Ar
                    ColumnText.ShowTextAligned(canvas2,
                          Element.ALIGN_RIGHT, new Phrase(ptRequiredAr, f2), 413, erListIndexAr, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                    erListIndexEn = erListIndexEn - 10;
                    erListIndexAr = erListIndexAr - 10;
                }


                //Extra Entry Requirements Pre requiste subject
                if (absApply.ExtraEntryRequirements == true)
                {

                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase("Need Prerequisite Subjects.", f3), 190, erListIndexEn, 0);

                    ColumnText.ShowTextAligned(canvas2,
                          Element.ALIGN_RIGHT, new Phrase("دراسة مواد تكميلية.", f2), 413, erListIndexAr, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                    erListIndexEn = erListIndexEn - 10;
                    erListIndexAr = erListIndexAr - 10;
                }

                //Structure A proposal
                if (absApply.Course1.CourseDescription.CourseLevel == "M" || absApply.Course1.CourseDescription.CourseLevel == "P")
                {
                    if (absApply.Course1.Code.Split('-')[1].Substring(0, 1) == "A")
                    {

                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase("Submit Thesis Proposal Required.", f3), 190, erListIndexEn, 0);

                        ColumnText.ShowTextAligned(canvas2,
                              Element.ALIGN_RIGHT, new Phrase("تقديم خطة البحث.", f2), 413, erListIndexAr, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                        erListIndexEn = erListIndexEn - 10;
                        erListIndexAr = erListIndexAr - 10;
                    }
                }

                if (erListIndexAr == 585 && erListIndexEn == 604)
                {


                    ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase("N/A", f3), 190, erListIndexEn, 0);
                    ColumnText.ShowTextAligned(canvas2,
                              Element.ALIGN_RIGHT, new Phrase("N/A", f2), 413, erListIndexAr, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                }

                //All fees

                int cell = 595;
                int i = 0;
                decimal totalMyr = 0m;
                string totalFee = "RM " + totalMyr + " Malaysian Ringgit";
                string totalFeeAr = totalMyr + " رنجيت ماليزي ";

                if (studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY")
                {
                    if (applicantFeeStructureItems != null)
                        foreach (var fs in applicantFeeStructureItems.OrderBy(c => c.Name))
                        {

                            ++i;
                            totalMyr = totalMyr + Convert.ToDecimal(fs.AmountMYR);
                            ColumnText.ShowTextAligned(canvas3, Element.ALIGN_LEFT, new Phrase(i.ToString() + ". ", f3), 77, cell, 0);
                            ColumnText.ShowTextAligned(canvas3, Element.ALIGN_LEFT, new Phrase("(" + fs.NameAr.ToString() + ") " + fs.Name.ToString(), f3), 105, cell, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            ColumnText.ShowTextAligned(canvas3, Element.ALIGN_LEFT, new Phrase(fs.AmountMYR.ToString("0.00"), f3), 450, cell, 0);

                            cell = cell - 15;
                        }
                    // Add immigration bond, calling visa and insurance
                    if (apply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONCAMPUS")
                    {

                        if (!absApply.PlacementTestExamRequired)
                        {
                            var itemToRemove = othersFees.Where(c => c.Name.Contains("MAPT/MEPT Fee")).LastOrDefault();
                            othersFees.Remove(itemToRemove);
                        }

                        foreach (var fee in othersFees)
                        {

                            ++i;
                            ColumnText.ShowTextAligned(canvas3, Element.ALIGN_LEFT, new Phrase(i.ToString() + ". ", f3), 77, cell, 0);
                            ColumnText.ShowTextAligned(canvas3, Element.ALIGN_LEFT, new Phrase("(" + fee.NameAr.ToString() + ") " + fee.Name.ToString(), f3), 105, cell, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            ColumnText.ShowTextAligned(canvas3, Element.ALIGN_LEFT, new Phrase(fee.AmountMYR.ToString("0.00"), f3), 450, cell, 0);
                            totalMyr = totalMyr + Convert.ToDecimal(fee.AmountMYR);

                            cell = cell - 15;
                        }

                    }

                    //TOTAL FEES
                    ColumnText.ShowTextAligned(canvas3, Element.ALIGN_LEFT, new Phrase(totalMyr.ToString("0.00"), f4), 450, cell, 0);
                    if ((studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY") && totalMyr < 5000)
                    {

                        totalFee = "Five thousand (RM 5000) Malaysian Ringgit.";
                        totalFeeAr = " خمسة آلاف (5000) رنجيت ماليزي ";

                        string balEn = "*Balance will calculated in the student’s payment based on the invoices issued.";
                        string balAr = "* " + " المتبقي من المبلغ المقدم " + "يتم احتسابه ضمن مدفوعات الطالب حسب ما يصدر له من فواتير  ";


                        // Balance
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(balEn, f4e), 63, 485, 0);

                        ColumnText.ShowTextAligned(canvas2,
                              Element.ALIGN_RIGHT, new Phrase(balAr, f4a), 402, 495, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                    }
                    else
                    {
                        totalFee = "RM " + totalMyr.ToString("0.00") + " Malaysian Ringgit";
                        totalFeeAr = totalMyr.ToString("0.00") + " رنجيت ماليزي ";
                    }

                    // 10 Total Fee
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(totalFee, f4e), 184, 506, 0);

                    // page 3 serial
                    ColumnText.ShowTextAligned(canvas3,
                            Element.ALIGN_LEFT, new Phrase(serialNumber, f2), 450, 800, 0);
                }






                // if (studyMode.ToUpper() == "ONLINE" && absApply.Applicant.Profile.Citizenship != "MY")
                //Stopped based on CEO request on 27/9/2016
                //    if (studyMode.ToUpper() == "ONLINE")
                //{

                //    totalFee = "Three thousand (RM 3000) Malaysian Ringgit.";
                //    totalFeeAr = " ثلاثة آلاف (3000) رنجيت ماليزي ";

                //    string balEn = "*Balance will calculated in the student’s payment based on the invoices issued.";
                //    string balAr = "* " + " المتبقي من المبلغ المقدم " + "يتم احتسابه ضمن مدفوعات الطالب حسب ما يصدر له من فواتير  ";


                //    // Balance
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(balEn, f4e), 63, 485, 0);

                //    ColumnText.ShowTextAligned(canvas2,
                //          Element.ALIGN_RIGHT, new Phrase(balAr, f4a), 402, 495, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                //}



                //TnC

                //if (studyMode.ToUpper() == "ONLINE" && absApply.Applicant.Profile.Citizenship != "MY")
                //{
                //    string tnc = "* This Preliminary admission offer letter is not valid for requesting student visa.";
                //    string tncar = "* هذا القبول المبدئي غير صالح لطلب استخراج تأشيرة الطالب للتعليم المباشر. ";

                //    ColumnText.ShowTextAligned(canvas,
                //         Element.ALIGN_LEFT, new Phrase(tnc, f4a), 83, 208, 0);

                //    ColumnText.ShowTextAligned(canvas2,
                //              Element.ALIGN_RIGHT, new Phrase(tncar, f4a), 517, 201, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                //}

                ////image sign picture Stamp
                canvas.AddImage(image);

                //Arabic -----------------------------------------------------------------------------------------------------------------------               


                // 1 Serial Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f2), 450, 800, 0);


                //A Learning Mode OL / OC
                //oncampus - التعليم المباشر
                //online - التعليم عن بعد
                string studyModeAr = "( التعليم عن بعد )";
                if (studyMode.ToUpper() == "ONCAMPUS")
                {
                    studyModeAr = "( التعليم المباشر )";
                }
                else
                {
                    studyModeAr = "( التعليم عن بعد )";
                }

                String arabicSemesterName = "";

                if (absApply.Course1.CourseDescription.CourseLevel == "P" && absApply.StartMonth != null)
                {
                    arabicSemesterName = " إشعار قبول مبدئي لفصل " + monthAr + " " + yearAr + " " + studyModeAr; ;
                }
                else
                {
                    arabicSemesterName = " إشعار قبول مبدئي لفصل " + semester.MonthNameAr + " " + semester.Year.ToString() + " " + studyModeAr;
                }
                // 5 Semester
                float seTextLength = canvas2.GetEffectiveStringWidth(arabicSemesterName, true);

                float position = ((60 - seTextLength) / 2) + 270;
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_CENTER, new Phrase(arabicSemesterName, f4), 299, 728, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 2 Reference Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.RefNo, f2), 413, 695, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 3 Name
                string ArabicName = absApply.Applicant.Profile.NameAr;
                if (String.IsNullOrEmpty(ArabicName) || ArabicName == "None")
                {
                    ArabicName = absApply.Applicant.Profile.Name;
                }

                float textLength = canvas2.GetEffectiveStringWidth(ArabicName, true);

                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(ArabicName, f2), 413, 681, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4 Date
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(currentDateTime.ToString("dd MMMMMM yyyy", new CultureInfo("ar-QA")), f2), 413, 667, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4.5 Passport Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(cardnum, f2), 413, 653, 0);

                // 6 Nationality
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(nationality.NameAr, f2), 413, 639, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 6 Program
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.Course1.NameAr, f2), 413, 625, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f2), 413, 611, 0);

                // 8.1 Normal type
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(Convert.ToString(durationTypeAr), f2), 403, 611, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 9 Vertual Account Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.HSBCAccPayor, f2), 413, 597, 0);// open on 14 August 2017

                // 10 Total Fee
                if (studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY")
                {
                    ColumnText.ShowTextAligned(canvas2,
                          Element.ALIGN_RIGHT, new Phrase(totalFeeAr, f4a), 180, 510, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                }

                ////image sign picture Stamp
                canvas2.AddImage(imageAr);

                stamper.Close();

            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// If dateOffterLetterSent is null, means that it will use the DateTime.Now()
        /// </summary>
        /// <param name="apply"></param>
        /// <param name="offerLetterFileName"></param>
        /// <param name="dateOfferLetterSent"></param>
        /// <returns></returns>
        public string GenerateStudentOfferLetter(AdmissionApply apply, string offerLetterFileName, DateTime? dateOfferLetterSent, string blankPrint)
        {
            string destOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PreviewOfferLetters\\" + offerLetterFileName;
            //string destOffer =  @"c:\cms_temp\OfferLetter\" + offerLetterFileName;
            // testing

            //GenerateUndergraduateFinalOffer(destOffer, apply, offerLetterFileName, dateOfferLetterSent);

            if (apply.Course1.CourseDescription.CourseLevel != "S")
            {
                // GenerateUniversialOfferLetter(destOffer, apply, offerLetterFileName, dateOfferLetterSent);
                GenerateOfficialOfferLetter(destOffer, apply, offerLetterFileName, dateOfferLetterSent, blankPrint);
            }

            else
            {
                GenerateShortCourseOfferLetter(destOffer, apply, offerLetterFileName, dateOfferLetterSent);

            }
            return destOffer;
        }

        public string GeneratePostgradOfficialLetter(PostGraduateThesis post, string offerLetterFileName, string type, bool isRinggit)
        {
            string destOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PreviewOfferLetters\\" + offerLetterFileName;
            //string destOffer =  @"c:\cms_temp\OfferLetter\" + offerLetterFileName;
            // testing
            if (type == "supervisor")
                GenerateSupervisorOfficialLetter(destOffer, post);
            else if (type == "chairman")
                GenerateChairmanOfficialLetter(destOffer, post);
            else if (type == "internalexaminer1" || type == "internalexaminer2")
                GenerateInternalExaminerOfficialLetter(destOffer, type, post);
            else if (type == "externalexaminer2" || type == "externalexaminer1")
                GenerateExternalExaminerOfficialLetter(destOffer, type, post, isRinggit);
            else if (type == "depsfacultyrep")
                GenerateDepFacRepOfficialLetter(destOffer, post);
            else if (type == "studentnotice")
                GenerateStudentNoticeOfficialLetter(destOffer, post);
            return destOffer;
        }

        public void GenerateSupervisorOfficialLetter(string destOffer, PostGraduateThesis post)
        {
            string sourceOffer = "";
            //Control Offer Letter Templete
            var student = post.Student;
            string mode = student.LearningModeCode;
            if (mode == "OC")
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\supervisor_appoint_OC.pdf";
            else
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\supervisor_appoint_OL.pdf";

            PdfReader r = new PdfReader(sourceOffer);
            using (var stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate)))
            {
                try
                {

                    string guid = Guid.NewGuid().ToString();

                    PdfContentByte canvas = stamper.GetOverContent(1);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);
                    //canvas2.SetFontAndSize(bf, 8);
                    //canvas3.SetFontAndSize(bf, 8);

                    Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f4 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    //Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    //Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                    //Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                    //Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                    //Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);
                    //English
                    var meetingMinutes = "";
                    var profile = student.Profile;
                    var course = student.Course;
                    var proposalrecord = post.ProposalApplication.ProposalRecordStatuses;
                    string nameEn = profile.NameEn;
                    string nameAr = profile.NameAr;
                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                    string SupervisorNameEn = post.SupervisorName;
                    string SupervisorNameAr = post.SupervisorNameAr;
                    // 1 Semester
                    string matrixNo = student.MatrixNo;
                    string refNo = student.UserName;
                    string courseNameEn = course.NameEn;
                    string CourseNameAr = course.NameAr;
                    var latestRecord = proposalrecord.Where(x => x.Status == ProposalStatus.ProposalAccepted).LastOrDefault();
                    string latestRecordreamrk = latestRecord.Remark;
                    string meetingDate = latestRecord.CreatedOn.ToString("dd/MM/yyyy");
                    const int minus10 = 10;
                    if (latestRecordreamrk.IndexOf(":") != -1)
                    {
                        meetingMinutes = latestRecordreamrk.Substring(0, latestRecordreamrk.IndexOf(':'));
                    }
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(SupervisorNameEn, f4), 97, 629, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(matrixNo, f2), 55, 563, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameEn, f2), 72, 575, 0);

                    var courseNameLine = BreakStringNewLine(courseNameEn);
                    if (courseNameLine.Count > 0)
                    {
                        int currentHeight = 538;
                        for (int i = 0; i < courseNameLine.Count; i++)
                        {
                            ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_LEFT, new Phrase(Convert.ToString(courseNameLine[i]), f2), 38, currentHeight, 0);
                            currentHeight -= minus10;
                        }

                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(courseNameEn, f2), 38, 538, 0);
                    }


                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f2), 100, 692, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(meetingMinutes, f2), 140, 506, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(meetingDate, f2), 60, 496, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(refNo, f3), 135, 701, 0);


                    //Arabic -----------------------------------------------------------------------------------------------------------------------               
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(SupervisorNameAr, f4), 420, 629, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(matrixNo, f2), 460, 512, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_RIGHT, new Phrase(nameAr, f2), 520, 525, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_RIGHT, new Phrase(CourseNameAr, f2), 534, 493, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f2), 475, 690, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingMinutes, f2), 425, 553, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingDate, f2), 490, 538, 0);
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(refNo, f3), 483, 702, 0);
                }
                catch
                {

                }
                finally
                {
                    stamper.SetFullCompression();
                    stamper.Writer.CompressionLevel = PdfStream.BEST_COMPRESSION;
                    stamper.Close();
                }
            }
        }

        /// <summary>
        /// Faris : For Internal examiner 1 or 2
        /// </summary>
        /// <param name="destOffer">path temp</param>
        /// <param name="type">internal only</param>
        /// <param name="post">postgrad object</param>
        public void GenerateInternalExaminerOfficialLetter(string destOffer, string type, PostGraduateThesis post)
        {
            string sourceOffer = "";

            //Control Offer Letter Templete
            var student = post.Student;
            string level = student.Level;
            string mode = student.LearningModeCode;
            if (level == "M") //master dude
            {
                if (student.DepsStructure == "A")
                {
                    if (mode == "OC")
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\internal_examiner_master_A_OC.pdf";
                    else
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\internal_examiner_master_A_OL.pdf";
                }
                else if (student.DepsStructure == "B")
                {
                    if (mode == "OC")
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\internal_examiner_master_B_OC.pdf";
                    else
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\internal_examiner_master_B_OL.pdf";
                }
                else if (student.DepsStructure == "C") // will implement future
                {

                }
                else
                {
                    throw new Exception("Not Recognize DEPS Structure");
                }

            }
            else if (level == "P") //PHD 
            {
                if (mode == "OC")
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\internal_examiner_PHD_OC.pdf";
                else
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\internal_examiner_PHD_OL.pdf";
            }
            else // get the hell out of here bachelor :D
            {
                throw new Exception("Not Elligible To Do This!");
            }


            PdfReader r = new PdfReader(sourceOffer);
            using (var stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate)))
            {
                try
                {

                    string guid = Guid.NewGuid().ToString();

                    PdfContentByte canvas = stamper.GetOverContent(1);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);
                    //canvas2.SetFontAndSize(bf, 8);
                    //canvas3.SetFontAndSize(bf, 8);

                    Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f4 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    Font f5 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f6 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    //Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    //Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                    //Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                    //Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                    //Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);
                    //English

                    var meetingMinutes = "";
                    var profile = student.Profile;
                    var course = student.Course;
                    string nameEn = profile.NameEn;
                    string nameAr = profile.NameAr;
                    string ExaminerNameEn = string.Empty;
                    string ExaminerNameAr = string.Empty;
                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                    string yearCode = currentDate.Substring(currentDate.Length - 2);
                    string monthCode = currentDate.Substring(3, 2);
                    var thesisreport = post.ThesisReports.Where(x => x.IsApplyingForFinalReport).LastOrDefault();
                    var finalreport = thesisreport.FinalReport;
                    if (type == "internalexaminer1")
                    {
                        var examinerprofile = finalreport.IpsInternalExaminer1.Profile;
                        ExaminerNameEn = examinerprofile.Title + examinerprofile.Name;
                        ExaminerNameAr = examinerprofile.NameAr;
                    }
                    else if (type == "internalexaminer2")
                    {
                        var examinerprofile = finalreport.IpsInternalExaminer2.Profile;
                        ExaminerNameEn = examinerprofile.Title + examinerprofile.Name;
                        ExaminerNameAr = examinerprofile.NameAr;
                    }
                    var finalreportrecord = thesisreport.ThesisReportRecords;
                    string matrixNo = student.MatrixNo;
                    string refNo = yearCode + "/" + svcThesisReportRunningNo.BuildRefCodeNumber();
                    string courseNameEn = course.NameEn;
                    string CourseNameAr = course.NameAr;
                    var latestRecord = finalreportrecord.Where(x => x.ThesisReportStateStatus.ThesisReportStatus == ThesisReportStatus.FinalReportApproved).LastOrDefault();
                    string latestRecordreamrk = latestRecord.Remark;
                    string meetingDate = latestRecord.CreatedOn.ToString("dd/MM/yyyy");
                    if (latestRecordreamrk.IndexOf(":") != -1)
                    {
                        meetingMinutes = latestRecordreamrk.Substring(0, latestRecordreamrk.IndexOf(':'));
                    }

                    //handle research title
                    string researchTitle = String.Empty;
                    //set max one line only 65 char
                    var researchtitlebreak = Regex.Split(post.ResearchTitle, @"(.{1,65})(?:\s|$)|(.{65})").Where(x => x.Length > 0).ToList();
                    var linecount = researchtitlebreak.Count;
                    //end handle reasearch title

                    //english-------------------------------------------------------------
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(refNo, f3), 150, 724, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 103, 713, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(ExaminerNameEn, f4), 116, 650, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(meetingMinutes, f2), 168, 588, 0);
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(meetingDate, f2), 86, 568, 0);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_LEFT, new Phrase(matrixNo, f2), 55, 563, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameEn, f2), 98, 557, 0);
                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 538; //starting for english
                        foreach (var item in researchtitlebreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 55, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 55, 538, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    var CoursenameBreak = Regex.Split(courseNameEn, @"(.{1,65})(?:\s|$)|(.{65})").Where(x => x.Length > 0).ToList();
                    var courseLinecount = CoursenameBreak.Count;
                    if (courseLinecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 490; //starting for english
                        foreach (var item in CoursenameBreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 50, currentHeight, 0);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(courseNameEn, f5), 50, 490, 0);
                    }

                    //Arabic -----------------------------------------------------------------------------------------------------------------------               
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(refNo, f3), 488, 724, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 475, 712, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(ExaminerNameAr, f5), 350, 649, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingMinutes, f2), 440, 585, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingDate, f2), 470, 570, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_RIGHT, new Phrase(nameAr, f5), 520, 558, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_MIDDLE, new Phrase(matrixNo, f2), 465, 508, 0);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_RIGHT, new Phrase(nameAr, f2), 520, 525, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG); 

                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 535; //starting for ar
                        foreach (var item in researchtitlebreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 360, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 360, 535, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }
                    var CoursenameArBreak = Regex.Split(CourseNameAr, @"(.{1,65})(?:\s|$)|(.{65})").Where(x => x.Length > 0).ToList();
                    var courseLineArcount = CoursenameArBreak.Count;
                    if (courseLineArcount > 1) // means the title too long dude 
                    {
                        int currentHeight = 485; //starting for english
                        foreach (var item in CoursenameArBreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_RIGHT, new Phrase(item, f5), 550, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas,
                             Element.ALIGN_RIGHT, new Phrase(CourseNameAr, f5), 550, 485, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }


                    stamper.SetFullCompression();
                    stamper.Writer.CompressionLevel = PdfStream.BEST_COMPRESSION;
                    stamper.FormFlattening = true;
                }
                catch
                {

                }
                finally
                {

                    stamper.Close();
                    r.Close();
                }
            }
        }

        /// <summary>
        /// Faris : For External examiner 1 or 2
        /// </summary>
        /// <param name="destOffer">path temp</param>
        /// <param name="type">xnternal only</param>
        /// <param name="post">postgrad object</param>
        public void GenerateExternalExaminerOfficialLetter(string destOffer, string type, PostGraduateThesis post, bool IsRinggit)
        {
            string sourceOffer = "";

            //Control Offer Letter Templete
            var student = post.Student;
            string level = student.Level;
            string mode = student.LearningModeCode;
            string Amount = string.Empty;
            if (level == "M") //master dude
            {
                if (student.DepsStructure == "A")
                {
                    if (mode == "OC")
                    {
                        if (IsRinggit)
                        {
                            Amount = "RM 500";
                            sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_master_A_MY_OC.pdf";
                        }
                        else
                        {
                            Amount = "EGP 750";
                            sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_master_A_EG_OC.pdf";
                        }
                    }
                    else
                    {
                        if (IsRinggit)
                        {
                            Amount = "RM 500";
                            sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_master_A_MY_OL.pdf";
                        }
                        else
                        {
                            Amount = "EGP 750";
                            sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_master_A_EG_OL.pdf";
                        }
                    }
                }
                else
                {
                    throw new Exception("Master External Only For Structure A!");
                }

            }
            else if (level == "P") //PHD 
            {
                if (mode == "OC")
                {
                    if (IsRinggit)
                    {
                        Amount = "RM 750";
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_PHD_MY_OC.pdf";
                    }
                    else
                    {
                        Amount = "EGP 1000";
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_PHD_EG_OC.pdf";
                    }
                }
                else
                {
                    if (IsRinggit)
                    {
                        Amount = "RM 750";
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_PHD_MY_OL.pdf";
                    }
                    else
                    {
                        Amount = "EGP 1000";
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\external_examiner_PHD_EG_OL.pdf";
                    }
                }
            }
            else // get the hell out of here bachelor :D
            {
                throw new Exception("Not Elligible To Do This!");
            }


            PdfReader r = new PdfReader(sourceOffer);
            using (var stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate)))
            {
                try
                {

                    string guid = Guid.NewGuid().ToString();

                    PdfContentByte canvas = stamper.GetOverContent(1);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);
                    //canvas2.SetFontAndSize(bf, 8);
                    //canvas3.SetFontAndSize(bf, 8);
                    Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f4 = new Font(bf, 9, Font.NORMAL, BaseColor.BLACK);
                    Font f5 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f6 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    //Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    //Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                    //Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                    //Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                    //Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);
                    //English
                    string examinertype = string.Empty;
                    examinertype = type.Substring(type.Length - 1);
                    var meetingMinutes = string.Empty;
                    var profile = student.Profile;
                    var course = student.Course;
                    string nameEn = profile.NameEn;
                    string nameAr = profile.NameAr;
                    string ExaminerNameEn = string.Empty;
                    string ExaminerNameAr = string.Empty;
                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                    string yearCode = currentDate.Substring(currentDate.Length - 2);
                    string monthCode = currentDate.Substring(3, 2);
                    var thesisreport = post.ThesisReports.Where(x => x.IsApplyingForFinalReport).LastOrDefault();
                    var finalreport = thesisreport.FinalReport;
                    if (type == "externalexaminer1")
                    {
                        var examinerprofile = finalreport.IpsExternalExaminer1.Profile;
                        ExaminerNameEn = examinerprofile.Title + " " + examinerprofile.Name;
                        ExaminerNameAr = examinerprofile.NameAr;
                    }
                    else if (type == "externalexaminer2")
                    {
                        var examinerprofile = finalreport.IpsExternalExaminer2.Profile;
                        ExaminerNameEn = examinerprofile.Title + " " + examinerprofile.Name;
                        ExaminerNameAr = examinerprofile.NameAr;
                    }
                    var finalreportrecord = thesisreport.ThesisReportRecords;
                    string matrixNo = student.MatrixNo;
                    string refNo = yearCode + "/" + svcThesisReportRunningNo.BuildRefCodeNumber();
                    string courseNameEn = course.NameEn;
                    string CourseNameAr = course.NameAr;
                    var latestRecord = finalreportrecord.Where(x => x.ThesisReportStateStatus.ThesisReportStatus == ThesisReportStatus.FinalReportApproved).LastOrDefault();
                    string latestRecordreamrk = latestRecord.Remark;
                    string meetingDate = latestRecord.CreatedOn.ToString("dd/MM/yyyy");
                    if (latestRecordreamrk.IndexOf(":") != -1)
                    {
                        meetingMinutes = latestRecordreamrk.Substring(0, latestRecordreamrk.IndexOf(':'));
                    }

                    //handle research title
                    string researchTitle = String.Empty;
                    //set max one line only 58 char
                    var researchtitlebreak = Regex.Split(post.ResearchTitle, @"(.{1,45})(?:\s|$)|(.{45})").Where(x => x.Length > 0).ToList();
                    var linecount = researchtitlebreak.Count;
                    //end handle reasearch title

                    //english-------------------------------------------------------------
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(refNo, f3), 156, 730, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 110, 716, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(ExaminerNameEn, f4), 47, 650, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(examinertype, f4), 237, 600, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(meetingMinutes, f2), 124, 587, 0);
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(meetingDate, f2), 80, 576, 0);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_LEFT, new Phrase(matrixNo, f2), 55, 563, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameEn, f2), 90, 565, 0);

                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 544; //starting for english
                        foreach (var item in researchtitlebreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 46, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 46, 544, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    var CoursenameBreak = Regex.Split(courseNameEn, @"(.{1,65})(?:\s|$)|(.{65})").Where(x => x.Length > 0).ToList();
                    var courseLinecount = CoursenameBreak.Count;
                    if (courseLinecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 496; //starting for english
                        foreach (var item in CoursenameBreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 45, currentHeight, 0);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(courseNameEn, f5), 45, 496, 0);
                    }

                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(Amount, f5), 193, 460, 0);
                    //Arabic -----------------------------------------------------------------------------------------------------------------------               
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(refNo, f3), 523, 731, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 510, 718, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(ExaminerNameAr, f5), 388, 650, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(examinertype, f4), 463, 590, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingMinutes, f2), 428, 580, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingDate, f2), 450, 567, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_RIGHT, new Phrase(nameAr, f5), 500, 554, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_MIDDLE, new Phrase(matrixNo, f2), 465, 508, 0);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_RIGHT, new Phrase(nameAr, f2), 520, 525, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG); 

                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 530; //starting for ar
                        foreach (var item in researchtitlebreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 310, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 310, 530, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(CourseNameAr, f5), 490, 491, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(Amount, f5), 392, 464, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    stamper.SetFullCompression();
                    stamper.Writer.CompressionLevel = PdfStream.BEST_COMPRESSION;
                    stamper.FormFlattening = true;
                }
                catch
                {

                }
                finally
                {

                    stamper.Close();
                    r.Close();
                }
            }
        }

        public void GenerateDepFacRepOfficialLetter(string destOffer, PostGraduateThesis post)
        {
            string sourceOffer = "";

            //Control Offer Letter Templete
            var student = post.Student;
            string mode = student.LearningModeCode;
            if (mode == "OC")
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\faculty_representative_OC.pdf";
            else
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\faculty_representative_OL.pdf";

            PdfReader r = new PdfReader(sourceOffer);
            using (var stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate)))
            {
                try
                {
                    string guid = Guid.NewGuid().ToString();
                    PdfContentByte canvas = stamper.GetOverContent(1);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);
                    //canvas2.SetFontAndSize(bf, 8);
                    //canvas3.SetFontAndSize(bf, 8);

                    Font f2 = new Font(bf, 9, Font.NORMAL, BaseColor.BLACK);
                    Font f3 = new Font(bf, 9, Font.NORMAL, BaseColor.BLACK);
                    Font f4 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f5 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f6 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f7 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    //Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    //Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                    //Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                    //Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                    //Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);
                    //English

                    var meetingMinutes = "";
                    var profile = student.Profile;
                    var course = student.Course;
                    string nameEn = profile.NameEn;
                    string nameAr = profile.NameAr;

                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                    string yearCode = currentDate.Substring(currentDate.Length - 2);
                    string monthCode = currentDate.Substring(3, 2);
                    var thesisreport = post.ThesisReports.Where(x => x.IsApplyingForFinalReport).LastOrDefault();
                    var finalreport = thesisreport.FinalReport;
                    var chair_profile = finalreport.DeptFacRepresentative.Profile;
                    string deprepEn = chair_profile.Title + chair_profile.Name;
                    string deprepAr = chair_profile.NameAr;
                    var finalreportrecord = thesisreport.ThesisReportRecords;
                    string matrixNo = student.MatrixNo;
                    string refNo = yearCode + "/" + svcThesisReportRunningNo.BuildRefCodeNumber();
                    string courseNameEn = course.NameEn;
                    string CourseNameAr = course.NameAr;
                    var latestRecord = finalreportrecord.Where(x => x.ThesisReportStateStatus.ThesisReportStatus == ThesisReportStatus.FinalReportApproved).LastOrDefault();
                    string latestRecordreamrk = latestRecord.Remark;
                    string meetingDate = latestRecord.CreatedOn.ToString("dd/MM/yyyy");
                    if (latestRecordreamrk.IndexOf(":") != -1)
                    {
                        meetingMinutes = latestRecordreamrk.Substring(0, latestRecordreamrk.IndexOf(':'));
                    }

                    //english-------------------------------------------------------------

                    //handle research title
                    string researchTitle = String.Empty;
                    //set max one line only 65 char
                    var researchtitlebreak = Regex.Split(post.ResearchTitle, @"(.{1,45})(?:\s|$)|(.{45})").Where(x => x.Length > 0).ToList();
                    var linecount = researchtitlebreak.Count;
                    //end handle reasearch title

                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(refNo, f3), 150, 713, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 103, 704, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(deprepEn, f4), 50, 584, 0);
                    ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(meetingMinutes, f2), 50, 487, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(meetingDate, f2), 90, 473, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameEn, f2), 50, 432, 0);
                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 395; //starting for english
                        foreach (var item in researchtitlebreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 50, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 55, 395, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    //handle course Name too long 
                    string courseName = String.Empty;
                    //set max one line only 65 char
                    var coursenameEnBreak = Regex.Split(courseNameEn, @"(.{1,60})(?:\s|$)|(.{60})").Where(x => x.Length > 0).ToList();
                    var linecountCourseEn = coursenameEnBreak.Count;
                    //end handle reasearch title
                    if (linecountCourseEn > 1) // means the title too long dude 
                    {
                        int currentHeight = 340; //starting for english
                        foreach (var item in coursenameEnBreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 50, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        courseName = courseNameEn;
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(courseName, f5), 50, 340, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    //Arabic -----------------------------------------------------------------------------------------------------------------------               
                    ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_LEFT, new Phrase(refNo, f3), 485, 713, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 470, 703, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(deprepAr, f5), 480, 590, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingMinutes, f7), 306, 506, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingDate, f2), 459, 490, 0);


                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameAr, f5), 470, 462, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    string researchTitleAr = String.Empty;
                    //set max one line only 65 char
                    var researchtitleArbreak = Regex.Split(post.ResearchTitle, @"(.{1,45})(?:\s|$)|(.{45})").Where(x => x.Length > 0).ToList();
                    var linecountAr = researchtitlebreak.Count;
                    if (linecountAr > 1) // means the title too long dude 
                    {
                        int currentHeight = 430; //starting for ar
                        foreach (var item in researchtitleArbreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 350, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitleAr = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 350, 430, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }
                    //handle course Name too long 
                    string courseNameAr = String.Empty;
                    //set max one line only 65 char
                    var coursenameArBreak = Regex.Split(CourseNameAr, @"(.{1,45})(?:\s|$)|(.{45})").Where(x => x.Length > 0).ToList();
                    var linecountCourseAr = coursenameArBreak.Count;
                    //end handle reasearch title
                    if (linecountCourseAr > 1) // means the title too long dude 
                    {
                        int currentHeight = 380; //starting for english
                        foreach (var item in coursenameArBreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 440, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        courseNameAr = CourseNameAr;
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(courseNameAr, f5), 440, 380, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    stamper.SetFullCompression();
                    stamper.Writer.CompressionLevel = PdfStream.BEST_COMPRESSION;
                    stamper.FormFlattening = true;
                }
                catch
                {

                }
                finally
                {

                    stamper.Close();
                    r.Close();
                }
            }
        }

        public void GenerateStudentNoticeOfficialLetter(string destOffer, PostGraduateThesis post)
        {
            string sourceOffer = "";

            //Control Offer Letter Templete
            var student = post.Student;
            string mode = student.LearningModeCode;
            if (mode == "OC")
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\student_notice_OC.pdf";
            else
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\student_notice_OL.pdf";

            PdfReader r = new PdfReader(sourceOffer);
            using (var stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate)))
            {
                try
                {
                    string guid = Guid.NewGuid().ToString();
                    PdfContentByte canvas = stamper.GetOverContent(1);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);
                    //canvas2.SetFontAndSize(bf, 8);
                    //canvas3.SetFontAndSize(bf, 8);

                    Font f2 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f4 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    Font f5 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    Font f6 = new Font(bf, 7, Font.NORMAL, BaseColor.BLACK);
                    //Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    //Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                    //Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                    //Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                    //Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);
                    //English

                    var profile = student.Profile;
                    var course = student.Course;
                    string nameEn = profile.NameEn;
                    string nameAr = profile.NameAr;
                    string structure = student.DepsStructure;
                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                    string yearCode = currentDate.Substring(currentDate.Length - 2);
                    string monthCode = currentDate.Substring(3, 2);
                    string supervisorEn = post.SupervisorName;
                    string supervisorAr = post.SupervisorNameAr;
                    string matrixNo = student.MatrixNo;
                    string refNo = yearCode + "/" + svcThesisReportRunningNo.BuildRefCodeNumber();
                    string courseNameEn = course.NameEn;
                    string CourseNameAr = course.NameAr;
                    const int minus10 = 10;
                    //handle research title
                    string researchTitle = String.Empty;


                    //english-------------------------------------------------------------

                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(refNo, f3), 135, 718, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 103, 708, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameEn, f2), 135, 635, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(matrixNo, f4), 85, 624, 0);

                    //set max one line only 65 char
                    var researchtitlebreak = Regex.Split(post.ResearchTitle, @"(.{1,45})(?:\s|$)|(.{45})").Where(x => x.Length > 0).ToList();
                    var linecount = researchtitlebreak.Count;

                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 560; //starting for english
                        for (int i = 0; i < linecount; i++)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(Convert.ToString(researchtitlebreak[i]), f5), 77, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight -= minus10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 82, 560, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }
                    //end handle reasearch title

                    //course name break line
                    var courseCount = BreakStringNewLine(courseNameEn);
                    if (courseCount.Count > 0)
                    {
                        for (int i = 0; i < courseCount.Count; i++)
                        {
                            int currentHeight = 518;
                            ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(courseNameEn, f4), 73, currentHeight, 0);
                            currentHeight -= minus10;
                        }
                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_LEFT, new Phrase(courseNameEn, f4), 73, 518, 0);
                    }

                    //end course name break line

                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(structure, f4), 120, 489, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(supervisorEn, f2), 108, 477, 0);


                    // Arabic---------------------------------------------------------------------------------------------------------------------- -
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(refNo, f3), 486, 717, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 490, 707, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_RIGHT, new Phrase(nameAr, f2), 420, 635, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(matrixNo, f4), 460, 620, 0);

                    if (linecount > 1) // means the title too long dude 
                    {

                        int currentHeight = 570; //starting for ar
                        for (int i = 0; i < linecount; i++)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_RIGHT, new Phrase(Convert.ToString(researchtitlebreak[i]), f5), 530, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight -= minus10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_RIGHT, new Phrase(researchTitle, f5), 530, 570, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    var courseNameLine = BreakStringNewLine(CourseNameAr);
                    if (courseNameLine.Count > 0)
                    {
                        int currentHeight = 524;
                        ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_RIGHT, new Phrase(CourseNameAr, f5), 530, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                        currentHeight -= minus10; //getting low
                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(CourseNameAr, f5), 530, 524, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_RIGHT, new Phrase(supervisorAr, f2), 480, 493, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                    stamper.SetFullCompression();
                    stamper.Writer.CompressionLevel = PdfStream.BEST_COMPRESSION;
                    stamper.FormFlattening = true;
                }
                catch (Exception e)
                {

                }
                finally
                {

                    stamper.Close();
                    r.Close();
                }
            }
        }

        private static List<string> BreakStringNewLine(string content)
        {
            var contentList = Regex.Split(content, @"(.{1,65})(?:\s|$)|(.{65})").Where(x => x.Length > 0).ToList();
            return contentList;
        }

        public void GenerateChairmanOfficialLetter(string destOffer, PostGraduateThesis post)
        {
            string sourceOffer = "";

            //Control Offer Letter Templete
            var student = post.Student;
            string mode = student.LearningModeCode;
            if (mode == "OC")
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\chairman_appoint_OC.pdf";
            else
                sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\PostgraduateLetter\\chairman_appoint_OL.pdf";

            PdfReader r = new PdfReader(sourceOffer);
            using (var stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate)))
            {
                try
                {

                    string guid = Guid.NewGuid().ToString();

                    PdfContentByte canvas = stamper.GetOverContent(1);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);
                    //canvas2.SetFontAndSize(bf, 8);
                    //canvas3.SetFontAndSize(bf, 8);

                    Font f2 = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                    Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    Font f4 = new Font(bf, 9, Font.BOLD, BaseColor.BLACK);
                    Font f5 = new Font(bf, 9, Font.BOLD, BaseColor.BLACK);
                    Font f6 = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                    //Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                    //Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                    //Font f4a = new Font(bf, 7, Font.BOLD, BaseColor.BLACK);
                    //Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                    //Font f5 = new Font(bf, 10, Font.BOLDITALIC, BaseColor.BLACK);
                    //English

                    var meetingMinutes = "";
                    var profile = student.Profile;
                    var course = student.Course;
                    string nameEn = profile.NameEn;
                    string nameAr = profile.NameAr;

                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                    string yearCode = currentDate.Substring(currentDate.Length - 2);
                    string monthCode = currentDate.Substring(3, 2);
                    var thesisreport = post.ThesisReports.Where(x => x.IsApplyingForFinalReport).LastOrDefault();
                    var finalreport = thesisreport.FinalReport;
                    var chair_profile = finalreport.ChairmanCommittee.Profile;
                    string chairmanEn = chair_profile.Title + chair_profile.Name;
                    string chairmanAr = chair_profile.NameAr;
                    var finalreportrecord = thesisreport.ThesisReportRecords;
                    string matrixNo = student.MatrixNo;
                    string refNo = yearCode + "/" + svcThesisReportRunningNo.BuildRefCodeNumber();
                    string courseNameEn = course.NameEn;
                    string CourseNameAr = course.NameAr;
                    var latestRecord = finalreportrecord.Where(x => x.ThesisReportStateStatus.ThesisReportStatus == ThesisReportStatus.FinalReportApproved).LastOrDefault();
                    string latestRecordreamrk = latestRecord.Remark;
                    string meetingDate = latestRecord.CreatedOn.ToString("dd/MM/yyyy");
                    if (latestRecordreamrk.IndexOf(":") != -1)
                    {
                        meetingMinutes = latestRecordreamrk.Substring(0, latestRecordreamrk.IndexOf(':'));
                    }

                    //handle research title
                    string researchTitle = String.Empty;
                    //set max one line only 65 char
                    var researchtitlebreak = Regex.Split(post.ResearchTitle, @"(.{1,65})(?:\s|$)|(.{65})").Where(x => x.Length > 0).ToList();
                    var linecount = researchtitlebreak.Count;
                    //end handle reasearch title

                    //english-------------------------------------------------------------
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(chairmanEn, f4), 65, 604, 0);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_LEFT, new Phrase(matrixNo, f2), 55, 563, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameEn, f2), 75, 465, 0);
                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 440; //starting for english
                        foreach (var item in researchtitlebreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 55, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                              Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 55, 440, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 55, 440, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(courseNameEn, f4), 46, 400, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 103, 707, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(meetingMinutes, f2), 36, 508, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(meetingDate, f2), 70, 494, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(refNo, f3), 133, 717, 0);


                    //Arabic -----------------------------------------------------------------------------------------------------------------------               
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(chairmanAr, f5), 330, 605, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_MIDDLE, new Phrase(matrixNo, f2), 465, 508, 0);
                    //ColumnText.ShowTextAligned(canvas,
                    //      Element.ALIGN_RIGHT, new Phrase(nameAr, f2), 520, 525, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(currentDate, f3), 475, 706, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingMinutes, f2), 430, 509, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_MIDDLE, new Phrase(meetingDate, f2), 445, 492, 0);
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(nameAr, f5), 380, 475, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                    if (linecount > 1) // means the title too long dude 
                    {
                        int currentHeight = 445; //starting for ar
                        foreach (var item in researchtitlebreak)
                        {
                            ColumnText.ShowTextAligned(canvas,
                            Element.ALIGN_LEFT, new Phrase(item, f5), 310, currentHeight, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                            currentHeight = currentHeight - 10; //getting low
                        }
                    }
                    else
                    {
                        researchTitle = post.ResearchTitle;
                        ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(researchTitle, f5), 310, 445, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }

                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(CourseNameAr, f5), 520, 400, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(refNo, f3), 485, 716, 0);
                    stamper.SetFullCompression();
                    stamper.Writer.CompressionLevel = PdfStream.BEST_COMPRESSION;
                    stamper.FormFlattening = true;
                }
                catch
                {

                }
                finally
                {

                    stamper.Close();
                    r.Close();
                }
            }
        }

        public void GenerateShortCourseOfferLetter(string destOffer, AdmissionApply apply, string offerLetterFileName, DateTime? dateOfferLetterSent)
        {
            DateTime currentDateTime = (dateOfferLetterSent != null) ? (DateTime)dateOfferLetterSent : DateTime.Now;
            var s = sm.OpenSession();
            string sourceOffer;
            //Calculate admin fee
            decimal totalMyr = 0;
            List<FeeStructureItem> applicantFeeStructureItems = new List<FeeStructureItem>();
            var courseGroupItem = s.GetOne<CourseGroupItem>().Where(x => x.Course).IsEqualTo(apply.Course1)
                .AndHasChild(c => c.CourseGroup).Where(x => x.IsActive).IsEqualTo(true).EndChild().Execute();
            if (courseGroupItem != null)
            {
                var feeStructure = s.GetOne<FeeStructure>().Where(x => x.CourseGroup).IsEqualTo(courseGroupItem.CourseGroup).Execute();
                foreach (var item in feeStructure.FeeStructureItems.Where(x => x.Stage == Stage.One))
                {
                    if (apply.LearningMode == LearningMode.OnCampus)
                    {
                        if (apply.Applicant.Profile.Citizenship == "MY")
                        {
                            if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.Local || item.StudentType == StudentType.Both))
                            {
                                applicantFeeStructureItems.Add(item);
                            }
                        }
                        else
                        {
                            if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.International || item.StudentType == StudentType.Both))
                            {
                                applicantFeeStructureItems.Add(item);
                            }
                        }

                    }
                    else
                    {
                        if (item.Mode == Mode.Online || item.Mode == Mode.Both)
                        {
                            applicantFeeStructureItems.Add(item);
                        }
                    }
                }

                if (applicantFeeStructureItems != null)
                    foreach (var fs in applicantFeeStructureItems.OrderBy(c => c.Name))
                    {

                        totalMyr = totalMyr + Convert.ToDecimal(fs.AmountMYR);

                    }
            }

            if (!apply.Course1.IsTraining)
            {
                if (apply.LearningModeCode == "OL")
                {
                    //if (apply.Applicant.Profile.Citizenship != "MY")
                    //{
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Short Course  Offer Letter (OL)  for Sep  2011 auto.pdf";
                    //}
                    //else
                    //{
                    //    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Short Course  Offer Letter MY.pdf";
                    //}
                }
                else
                {

                    //if (apply.Applicant.Profile.Citizenship != "MY")
                    //{
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Short Course  Offer Letter (OC)  for Sep  2011 auto.pdf";
                    //}
                    //else
                    //{
                    //    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Short Course  Offer Letter MY.pdf";
                    //}
                }
            }
            else
            {
                if (apply.Applicant.AdmissionApply.Course1.Code == ("SCSIA"))// Arabic Short Course Mubeen
                {
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Training Offer Letter Arabic Course Summer.pdf";

                }
                else if (apply.Applicant.AdmissionApply.Course1.Code == ("SES")) // English Short Course Summer
                {
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Training  Offer Letter  fixed_English.pdf";
                }
                else
                {
                    if (totalMyr.Equals(0m))
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Training  Offer Letter  fixed.pdf";
                    else
                        sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Training _Offer_Letter_Pay.pdf";
                }
            }

            try
            {
                IntakeEvent semester = apply.Intake;
                var semesters = s.GetAll<CampusEvent>()
                                            .ThatHasChild(c => c.ParentEvent)
                                                .Where(c => c.Id).IsEqualTo(semester.ParentEvent.Id)
                                            .EndChild()
                                            .And(c => c.EventType).IsEqualTo("SME")
                                            .Execute().ToList();
                CampusEvent semesterDates = semesters.Where(c => c.Code.ToLower().StartsWith(semester.Code.Substring(0, 3).ToLower())).FirstOrDefault();

                AdmissionApply absApply = apply;

                int serial = semester.OfferLetterSerialNumber + 1;
                string cardnum = absApply.Applicant.Profile.IdCardNumber;
                if (cardnum == "nil")
                {
                    cardnum = "";
                }

                PdfReader r = new PdfReader(sourceOffer);
                string guid = Guid.NewGuid().ToString();
                PdfStamper stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate));
                PdfContentByte canvas = stamper.GetOverContent(1);
                PdfContentByte canvas2 = stamper.GetOverContent(2);
                PdfContentByte canvas3 = stamper.GetOverContent(3);
                BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                canvas.SetFontAndSize(bf, 8);
                canvas2.SetFontAndSize(bf, 8);

                Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                Font f5 = new Font(bf, 10, Font.ITALIC, BaseColor.BLACK);
                //English




                //string accNo = "50-009463-901";//R&D acc No
                //DateTime dateT = new DateTime();
                //if (!apply.Course1.IsTraining)
                //{
                //    accNo = absApply.HSBCAccPayor;
                //}

                string accNo = absApply.HSBCAccPayor;
                DateTime dateT = new DateTime();

                // 1 Semester
                //semester.Month


                // 1 Semester
                string semesterName = "SHORT COURSE ADMISSION FOR " + semester.MonthName.ToUpper() + " " + semester.Year.ToString() + " ( " + apply.LearningMode.ToString().ToUpper() + " )";
                if (apply.Course1.IsTraining)
                {
                    // semesterName = "TRAINING ADMISSION FOR " + dateT.Year.ToString("YYYY") + " ( " + apply.LearningMode.ToString().ToUpper() + " )"; edited by mohamed awadh for summer short course
                    semesterName = "TRAINING ADMISSION FOR " + absApply.Course1.NameEn.ToUpper();
                }

                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_CENTER, new Phrase(semesterName, f5), 300, 727, 0);

                //String semesterName = semester.MonthName.ToUpper() + " " + semester.Year.ToString();

                //ColumnText.ShowTextAligned(canvas,
                //      Element.ALIGN_CENTER, new Phrase(semesterName, f5), 310, 727, 0);

                // 2 Serial Number
                string serialNumber = serial.ToString("0000000") + "(" + currentDateTime.Year.ToString().Substring(2, 2) + currentDateTime.Month.ToString("00") + ")";
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);

                // 3 Reference Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.RefNo, f3), 204, 695, 0);

                // 4 Name
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.NameEnglish, f3), 204, 684, 0);

                // 5 Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(currentDateTime.ToString("dd.MM.yyyy"), f3), 204, 673, 0);

                // 6 Passport Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.IdCardNumber, f3), 204, 662, 0);

                // 7 Program
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Course1.NameEn, f3), 204, 652, 0);

                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f3), 204, 640, 0);

                var dur = absApply.Course1.CourseDescription.DurationType;
                string durationType = "";
                string durationTypeAr = "";
                if (dur == DurationType.Months)
                {
                    durationType = "Months";
                    durationTypeAr = "(أشهر)";
                }
                else if (dur == DurationType.Weeks)
                {
                    durationType = "Weeks";
                    durationTypeAr = "(أسابيع)";
                }
                else if (dur == DurationType.Days)
                {
                    durationType = "Days";
                    durationTypeAr = "(أيام)";
                }
                else
                {
                    durationType = "Years";
                    durationTypeAr = "(سنة/سنوات)";
                }
                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(durationType, f3), 215, 640, 0);

                if (!apply.Course1.IsTraining) //training dont show virtual account
                {
                    // 9 Vertual Account Number
                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(accNo, f3), 204, 629, 0);
                }
                //Regex rg;
                String ZipRegex = @"^SAS";


                //9 tuitionFee
                int tuitionFee = 0;
                //if (absApply.Applicant.Profile.Citizenship != "MY")
                //{
                if (!Regex.IsMatch(absApply.Course1.Code, ZipRegex))
                //if (absApply.Course1.Code.Contains("SAS"))
                {
                    tuitionFee = 2000;
                }
                else
                {

                    tuitionFee = 1500;
                }
                //}

                if (!apply.Course1.IsTraining)
                {


                    //if (apply.Applicant.Profile.Citizenship != "MY")
                    //{
                    ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_LEFT, new Phrase(tuitionFee.ToString("0"), f4e), 231, 469, 0);
                    //}

                    ColumnText.ShowTextAligned(canvas,
                          Element.ALIGN_LEFT, new Phrase(totalMyr.ToString("0"), f4e), 231, 457, 0);


                }
                else
                {
                    if (!totalMyr.Equals(0m))
                    {
                        ColumnText.ShowTextAligned(canvas,
                         Element.ALIGN_LEFT, new Phrase(totalMyr.ToString("0"), f4e), 231, 490, 0);
                        ColumnText.ShowTextAligned(canvas2,
                       Element.ALIGN_LEFT, new Phrase(totalMyr.ToString("0"), f4e), 328, 545, 0);
                    }
                }



                //if (apply.Intake.Code == ("APR 2016"))
                //{
                //    String arabicSemesterName = "TRAINING ADMISSION FOR " + absApply.Course1.NameEn.ToUpper();
                //    ColumnText.ShowTextAligned(canvas,
                //    Element.ALIGN_CENTER, new Phrase(arabicSemesterName, f4), 280, 728, 0);

                //    //String semesterName = semester.MonthName.ToUpper() + " " + semester.Year.ToString();

                //    //ColumnText.ShowTextAligned(canvas,
                //    //      Element.ALIGN_CENTER, new Phrase(semesterName, f5), 310, 727, 0);

                //    // 2 Serial Number
                //     serialNumber = serial.ToString("0000000") + "(" + currentDateTime.Year.ToString().Substring(2, 2) + currentDateTime.Month.ToString("00") + ")";
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);

                //    // 3 Reference Number
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(absApply.RefNo, f3), 204, 695, 0);

                //    // 4 Name
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.NameEnglish, f3), 204, 684, 0);

                //    // 5 Date
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(currentDateTime.ToString("dd.MM.yyyy"), f3), 204, 673, 0);

                //    // 6 Passport Number
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.IdCardNumber, f3), 204, 662, 0);

                //    // 7 Program
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(absApply.Course1.NameEn, f3), 204, 652, 0);

                //    // 8 Normal Period
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f3), 204, 640, 0);

                //     dur = absApply.Course1.CourseDescription.DurationType;
                //     durationType = "";
                //     durationTypeAr = "";
                //    if (dur == DurationType.Months)
                //    {
                //        durationType = "Months";
                //        durationTypeAr = "(أشهر)";
                //    }
                //    else if (dur == DurationType.Weeks)
                //    {
                //        durationType = "Weeks";
                //        durationTypeAr = "(أسابيع)";
                //    }
                //    else
                //    {
                //        durationType = "Years";
                //        durationTypeAr = "(سنة/سنوات)";
                //    }
                //    // 8 Normal Period
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(durationType, f3), 215, 640, 0);

                //    // 9 Vertual Account Number
                //    ColumnText.ShowTextAligned(canvas,
                //          Element.ALIGN_LEFT, new Phrase(accNo, f3), 204, 629, 0);

                //    //Regex rg;
                //     ZipRegex = @"^SAS";
                //}

                //else
                //{
                //Arabic                

                //A Learning Mode OL / OC
                //oncampus - التعليم المباشر
                // online - التعليم عن بعد
                string studyModeAr = "( التعليم عن بعد )";
                if (apply.LearningMode.ToString().ToUpper() == "ONCAMPUS")
                {
                    studyModeAr = "( التعليم المباشر )";
                }
                else
                {
                    studyModeAr = "( التعليم عن بعد )";
                }

                //ColumnText.ShowTextAligned(canvas2,
                //        Element.ALIGN_CENTER, new Phrase(studyModeAr, f4), 240, 728, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                String arabicSemesterName = " إشعار قبول مبدئي لفصل " + semester.Year.ToString() + " " + studyModeAr;

                if (apply.Course1.IsTraining)
                {
                    arabicSemesterName = " إشعار قبول مبدئي " + absApply.Course1.NameAr.ToUpper();
                }

                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_CENTER, new Phrase(arabicSemesterName, f4), 280, 728, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                // 1 Serial Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f2), 450, 800, 0);

                // 2 Reference Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.RefNo, f2), 413, 694, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 3 Name
                string ArabicName = absApply.Applicant.Profile.NameAr;
                if (String.IsNullOrEmpty(ArabicName) || ArabicName == "None")
                {
                    ArabicName = absApply.Applicant.Profile.Name;
                }

                float textLength = canvas2.GetEffectiveStringWidth(ArabicName, true);

                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(ArabicName, f2), 413, 680, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4 Date
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(currentDateTime.ToString("dd MMMMMM yyyy", new CultureInfo("ar-QA")), f2), 413, 666, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4.5 Passport Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(cardnum, f2), 413, 652, 0);

                //String arabicSemesterName = semester.MonthNameAr + " " + semester.Year.ToString();

                //// 5 Semester
                //float seTextLength = canvas2.GetEffectiveStringWidth(arabicSemesterName, true);

                //float position = ((60 - seTextLength) / 2) + 270;
                //ColumnText.ShowTextAligned(canvas2,
                //      Element.ALIGN_CENTER, new Phrase(arabicSemesterName, f4), 250, 728, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 6 Program
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.Course1.NameAr, f2), 413, 638, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f2), 413, 624, 0);

                // 8.1 Normal type
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(Convert.ToString(durationTypeAr), f2), 403, 624, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 9 tuitionFee
                if (!apply.Course1.IsTraining)
                {

                    //if (apply.Applicant.Profile.Citizenship != "MY")
                    //{
                    ColumnText.ShowTextAligned(canvas2,
                         Element.ALIGN_RIGHT, new Phrase(accNo, f2), 413, 611, 0);
                    ColumnText.ShowTextAligned(canvas2,
                        Element.ALIGN_LEFT, new Phrase(tuitionFee.ToString("0"), f4e), 332, 499, 0);
                    //}

                    ColumnText.ShowTextAligned(canvas2,
                          Element.ALIGN_LEFT, new Phrase(totalMyr.ToString("0"), f4e), 332, 485, 0);
                    //  }
                }

                stamper.Close();

            }
            catch (Exception ex)
            {

            }
        }

        public bool SendApplicationSubmittedEmailToApplicant(AdmissionApply apply, string verificationBaseUrl)
        {
            var subject = "Admission Application Submitted (" + apply.RefNo + ")";

            var svcApply = applyServiceFactory.GetService("ADM");

            var hash = svcApply.CreateVerificationHash(apply);

            //var hash = svcAdmissionApply.CreateVerificationHash(apply as AdmissionApply);
            var verifyUrl = verificationBaseUrl + "?r=" + apply.RefNo + "&h=" + hash;

            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <p>Name : " + apply.ApplicantName + @"</p>
                                <p>Reference Number : " + apply.RefNo + @"
                                <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                                <p>
                                    Thank you, your have successfully submitted your application under the reference number  <strong>" + apply.RefNo + @"</strong>.
                                    Please keep this number to communicate with us in future.
                                </p>

                                <p>
                                    Please note the following: To complete the registration process, you are required to confirm your registration by clicking on the following link</p>
                                    <a href=""" + verifyUrl + @""">" + verifyUrl + @"</a>
                                

                                <p>
                                    If you cannot open the above link, please contact the support team via the University e-mail address: <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a>. Please write your name and reference number clearly 
                                    or you can call telephone number +603 55113939 or mobile 0163516052 (Malaysia).
                                </p>

                                <p>
                                    After activating the code, you'll be able to follow up your application status and to upload all the required and necessary documents to support the application.
                                </p>
                                <p>Thank you! </p>
                                
                                Student Affairs Division<br>
                                Al-madinah International University
                                
                            </td>
                            <td width=""50%"" dir=""rtl"">
                                <p><b>الاسم:</b>
                                " + apply.ApplicantName + @"</p>
                                <p>الرقم المرجعي: 
                                " + apply.RefNo + @"</p>
                                <p>السلام عليكم ورحمة الله وبركاته</p>
                                <p>شكراً لك، طلبك قد اعتمد تحت رقم التسجيل المبين هنا
                                " + apply.RefNo + @". الرجاء الاحتفاظ بهذا الرقم للتواصل معنا مستقبلاً.
                                <p>يرجى ملاحظة ما يلي: لغرض إكمال عملية التسجيل، يجب عليك تأكيد تسجيل طلبك وذلك بالضغط على الرابط الآتي: </p>
                                <a href=""" + verifyUrl + @""">" + verifyUrl + @"</a>
                                <p>في حال عدم تمكنك من فتح الرابط أعلاه الرجاء الاتصال بفريق الدعم الخاص بالجامعة عبر عنوان البريد الإلكتروني التالي <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a>, (الرجاء كتابة الاسم والرقم المرجعي بوضوح) أو على رقم الهاتف 0060355113939 أو موبايل 0060163516052 (ماليزيا). </p>
                                <p>بعد تفعيل الكود، سيكون بإمكانك الدخول إلى نافذة ( المتقدمون للدراسة ) في جامعة المدينة العالمية لمتابعة الطلب الخاص بك ورفع وثائقك ومعرفة الرسوم الواجب تسديدها: 
                                
                                </p>
                                <p>شكراً لك!</p>
                                وكالة الشئون الطلابية<br>
                                جامعة المدينة العالمية          

                            </td>
                        </tr>
                    </table>
                    ";

            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public bool SendDocumentRequestEmailToApplicant(AdmissionApply apply, string applicantPortalUrl)
        {
            var subject = "Supporting Documents Needed (" + apply.RefNo + ")";

            var svcApply = applyServiceFactory.GetService("ADM");

            var password = apply.Applicant.Profile.GetPasswordPlainText();

            var body = @"
<table width=""100%"">
    <tr>
        <td width=""50%"" dir=""ltr"" valign=""top"">
            <p>Dear applicant</p>

            <p>Assalamualaikom warahmatullah wabarakatuh</p>

            <p>
            Thank you for submitting an application to enroll into Al-Madinah International University. Your application details are as follows: 
            
                <p style=""margin:20px""><b>
                Name: " + apply.Applicant.Profile.Name + @"<br />
                Reference Number: " + apply.RefNo + @"<br />
                Username: " + apply.Applicant.Profile.Username + @"<br />
                Password: " + password + @"
                </b></p>
            </p>

            <p>
            We hereby would like to draw your attention to the importance of the submission of all the required and necessary documents to support the application 
            during this week (true copies of academic certificates with the transcripts + copy of identity card + 3 passport-size photographs) in the following ways:
            </p>
            
            <p>
                <ol>
                    <li>Soft copy to the following link: <br />
                        <a href=""" + applicantPortalUrl + @""">" + applicantPortalUrl + @"</a><br />
                        Or through university website login into applicants (using the same username and password mentioned earlier)
                    </li>
                    
                    <li>Soft copy to the following e-mail: <br />
                        <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a></li>
                    
                    <li>By post to the following address:
                        <p style=""margin:5px 20px"">
                        Al-Madinah International University<br />
                        11th Floor, Plaza Masalam, <br />
                        No.2, Jalan Tengku Ampuan Zabedah E/9E, <br />
                        Section 9, 40100 Shah Alam, Selangor, Malaysia. <br />
                        TEL: +603 - 5511 3939 Ext: 404 Fax No: 03 - 5511 3940 
                        </p>
                    </li>
                </ol>
            </p>
            
            <p>
            If you encounter any problem in this regard, please contact the support team via the following channels:
            
                <ol>                
                    <li>Zopim Service“chatting”, from out website<br />                    
                        <a href=""http://www.mediu.edu.my/"">http://www.mediu.edu.my/</a> or;
                    </li>
                    
                    <li>Inquiry page<br /> 
                    <a href=""http://cmsweb.mediu.edu.my/EnquiryForm/EnquiryForm.aspx"">http://cmsweb.mediu.edu.my/EnquiryForm/EnquiryForm.aspx</a> or;
                    </li>
                    
                    <li>Via only SMS to ""0060193015078"" or;
                    </li>
                    
                    <li>Call telephone number +603 55113939 “Office hours 8:30am - 5:30pm Malaysia time - From Monday to Friday”
                    </li>
                </ol>
                
                For better service please always write your name and reference number clearly
            </p>

            <p>
            Thank you very much!<br /><br />
            <b>Students Recruitment Unit</b><br />
            Al-Madinah International University
            </p>
        </td>
        
        <td width=""50%"" dir=""rtl"" valign=""top"">
        
            <p>عزيزي المتقدم</p> 
            
            <p>السلام عليكم ورحمة الله وبركاته</p>
            
            <p>شكراً لكم، جامعة المدينة العالمية ترحب بك راغباً في الانضمام إلى ركبها ويسعدنا أن نبلغكم بأن طلب التقديم قد وصلنا  وتم اعتماده بالمعلومات التالية :</p>
            
            <p style=""margin:20px""><b>
            الإسم : " + apply.Applicant.Profile.Name + @"<br />
            الرقم المرجعي : " + apply.RefNo + @" <br />
            اسم المستخدم : " + apply.Applicant.Profile.Username + @" <br />
            كلمة المرور: " + password + @"
            </b></p>
            
            <p>
            ونود هنا أن نلفت انتباهكم إلى سرعة إرسال وثائقكم (نسخ طبق الأصل للشهادات الدراسية مع كشف الدرجات+صورة من بطاقة الهوية أو الجواز+3صور شخصية) خلال هذا الأسبوع وذلك عبر إحدى الطرق التالية:
            
                <ol>
                    <li>
                    سحبها بالماسح الضوئي- الاسكانر- وتحميلها إلى الرابط التالي:<br />
                    <a href=""http://online.mediu.edu.my/applicantportal/Home/Login"">http://online.mediu.edu.my/applicantportal/Home/Login</a><br />
                    أو عبر نافذة المتقدمين بموقع الجامعة الالكتروني
                    </li>
                    
                    <li>
                    سحبها بالماسح الضوئي –الاسكانر- وارسالها الى البريد الالكتروني التالي:</br>
                    <a href=""mailto:admission@mediu.edu.my"">admission@mediu.edu.my</a>
                    </li>
                    
                    <li>
                    إرسالها بالبريد السريع إلى العنوان التالي :
                        <p style=""margin:5px 20px"" dir=""ltr"">
                            Al-Madinah International University<br />
                            11th Floor, Plaza Masalam,<br />
                            No.2, Jalan Tengku Ampuan Zabedah E/9E,<br />
                            Section 9, 40100 Shah Alam, Selangor, Malaysia.<br />
                            TEL : +603 - 5511 3939   Ext : 720  Fax No : 03 - 5511 3940                  
                        </p>
                    </li>
                </ol>            
            </p>
            
            <p>
            وإذا كان يوجد لديكم أي استفسار بهذا الخصوص، الرجاء الاتصال بفريق الدعم الخاص بالجامعة عبر الطرق التالية: 
            
                <ol>
                    <li>
                        خدمة المساعدة المباشرة<br />
                        <span dir=""ltr"">Zopim</span><br />
                        ""اضغط على الرابط أدناه وسيظهر لك اطار منبثق يمكنكم استخدامه للتحدث مباشرة مع موظفي الجامعة"" أو<br />
                        <a href=""http://www.mediu.edu.my/?lang=ar"">http://www.mediu.edu.my/?lang=ar</a>
                    </li>
                    
                    <li>
                    خدمة الاستفسار ""اضغط على الرابط أدناه"" <br />
                    <a href=""http://cmsweb.mediu.edu.my/EnquiryForm/EnquiryForm.aspx?lang=ar-SA"">http://cmsweb.mediu.edu.my/EnquiryForm/EnquiryForm.aspx?lang=ar-SA</a>
                    </li>
                    
                    <li>
                    أو أرسل رسالة قصيرة إلى الهاتف الجوال ""0060193015078""(لخدمة أفضل الرجاء كتابة الاسم والرقم المرجعي)
                    </li>
                    
                    <li>
                        أو الاتصال مباشرة على رقم الهاتف<br />
                        <span dir=""ltr"">+603 55113939</span><br />
                        (وقت الدوام فقط من 8:30 صباحاً  حتى 5:30 مساء بتوقيت ماليزيا"" من يوم الاثنين إلى يوم الجمعة)
                    </li>
                </ol>

            </p>
            <p>
            شكراً لك!<br /><br />
            <b>وحدة استقطاب الطلاب</b><br />
            جامعة المدينة العالمية
            </p>        
        </td>
    </tr>
</table>";

            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public bool SendMasterAttendanceFirstWarningLetter(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = subjectRegistered.TutorialGroup.Tutor.Profile.Email;
            }

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "Academic Warning (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>Institute of Postgraduate studies</h3></center>
                            <center><h3>ACADEMIC WARNING</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            Department: Institute of Postgraduate Studies<br />
                            <p><b>Subject: Academic Warning (due to the Non-attendance of Direct Meetings)</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>Please be informed, that due to your treble absence in attending the direct meetings of the subject: (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") that I am lecturing, 
                            and due to the fact that there is no legitimate excuse submitted by you. I, hereby, give you this Academic Warning, 
                            coinciding with the Article No 29, and 30 of the Postgraduate rules, which stated that a minimum of 70% attendance 
                            for classes of the subjects is compulsory, and if candidate is unable to achieve this percentage he shall not be 
                            eligible to be admitted to the examination for any subject.</p>
                            
                            <p>We hope that you will adhere to your study attendance, and to obey the Academic policies.</p>
                            
                            <p>We pray to Allah Almighty to give you all the success and prosperity.</p>
                            
                            <br /><br /><br />
                            Lecturer of the subject<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>معهد الدراسات العليا</h3></center>
                            <center><h3>إنذار أكاديمي</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            القسم: Institute of Postgraduate Studies<br />
                            <p><b>الموضوع: إنذار أكاديمي بسبب عدم حضور اللقاءات المباشرة</b></p>
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            <p>أحيطكم علماً بأنه بسبب غيابكم عن حضور اللقاء المباشر في مادة (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") والتي أقوم بتدريسها، ولعدم تقديمكم عذراً مقبولاً لذلك، فإني أوجه لكم إنذاراً أكاديمياً، اعتبارا بما ورد في لائحة الدراسات العليا (المادة 29، والمادة 30)
                             واللتان تنصان على أنه يتعين على الطالب تحقيق نسبة حضور لا تقل عن 70% وإذا لم يتمكن من ذلك فإنه يكون معرضا للحرمان من الدخول إلى الامتحان النهائي.   
                            آملين منكم حسن متابعة دراستكم، والالتزام بالأنظمة الأكاديمية 
                            </p>
                            <p>داعين الله تعالى لك بالتوفيق والنجاح.</p>
                            <br /><br /><br />
                            
                            محاضر المادة<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    SaveEmailMessage(
                        subjectRegistered.Student, subjectRegistered.TutorialGroup.Tutor.Profile.Email,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "FIRST ATTENDANCE WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        public bool SendMasterAttendanceSecondWarningLetter(SubjectRegistered subjectRegistered, string emailTo)
        {
            throw new NotImplementedException();
        }

        public bool SendMasterAttendanceProhibitLetter(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            var dean = subjectRegistered.TutorialGroup.Tutor.Department.Faculty.Deanship;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = dean.Profile.Email;
            }
            //if (emailTo == "student")
            //{
            //    recipientEmail = subjectRegistered.Student.Profile.Email;
            //}
            //else
            //{
            //    recipientEmail = subjectRegistered.TutorialGroup.Tutor.Profile.Email;
            //}

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "Prohibition Attending Final Exam (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>PROHIBITION FROM ATTENDING THE FINAL EXAM</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            Department: Institute of Postgraduate Studies<br />
                            <p><b>Subject: Notification of Prohibition from Attending the Final Exam</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>I would like to draw your attention, that the percentage of your overall attendance of the total
                            practical lectures, and the overall participation in all Evaluated Activities of On-Line 
                            Teaching in every subject each semester is less than 70%. Hence in coinciding with the Article No: 14th of the Policy 
                            of Studying and Examination it is decided to prohibit you from attending the final exam of 
                            the subject (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @").</p>
                            
                            <br /><br /><br />
                            Dean of the faculty<br />
                            (" + dean.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>إشعار بالحرمان من دخول الامتحان النهائي</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            القسم: Institute of Postgraduate Studies<br />
                            <p><b>الموضوع: إشعار بالحرمان من دخول الامتحان النهائي</b></p>
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            
                            <p>أود أن أحيطكم علماً بأنه بسبب أن نسبة حضوركم قلت عن نسبة (70%) من مجموع حضور المحاضرات العامة والمحاضرات التطبيقية المحددة وأداء نشاطات التقييم المستمر في التعليم عن بعد؛ فعليه سيتم حرمانكم من دخول الامتحان النهائي.</p>
                            
                            <p>في مادة (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") بناءً على ما ورد في القاعدة التنفيذية للمادة الرابعة عشرة من نظام الدراسة والامتحانات.</p>
                            
                            <br /><br /><br />
                            
                            عميد الكلية<br />
                            (" + dean.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    string senderEmail = "";
                    if (dean != null)
                    {
                        senderEmail = dean.Profile.Email;
                    }
                    SaveEmailMessage(
                        subjectRegistered.Student, senderEmail,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "ATTENDANCE PROHIBIT WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        public bool SendMasterAttendanceFirstWarningLetterOnCampus(SubjectRegistered subjectRegistered, string emailTo)
        {
            throw new NotImplementedException();
        }

        public bool SendMasterAttendanceSecondWarningLetterOnCampus(SubjectRegistered subjectRegistered, string emailTo)
        {
            throw new NotImplementedException();
        }

        public bool SendMasterAttendanceProhibitLetterOnCampus(SubjectRegistered subjectRegistered, string emailTo)
        {
            throw new NotImplementedException();
        }

        public bool SendBachelorAttendanceFirstWarningLetter(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = subjectRegistered.TutorialGroup.Tutor.Profile.Email;
            }

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "First Academic Warning (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>THE FIRST ACADEMIC WARNING</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            
                            <p><b>Subject: The First Academic Warning due to the Non-attendance of Direct Meetings</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>Please be informed that due to your failure to attend the direct meetings of the subject (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") that 
                            I am lecturing for two times, and due to the fact that there is no legitimate excuse submitted by you, I, hereby, 
                            give you the first academic warning coinciding with the Article no.14 of the Studying and Examination 
                            Policy, which stated that the student who failed to attend the direct meetings twice must be given the 
                            First Academic warning.</p>
                            
                            <p>I am also bringing to your attention, that the above article defines that the student whose absence 
                            reached three times he will receive the Second Academic Warning.</p>
                            
                            <p>For your information, if the percentage of your attendance is less than 70% from the total attendance of 
                            the total practical lectures and the overall participation in all Evaluated 
                            Activities of On-Line Teaching in every subject each semester, the letter of Prohibition from Attending the 
                            Final Exam will be issued to you.</p>
                            
                            <p>We hope that you will adhere to your study attendance, and obey the academic policies.</p>
                            
                            <p>We pray to Allah Almighty to give you all the success and prosperity.</p>
                            
                            <br /><br /><br />
                            Lecturer of the subject<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>الإنذار الأكاديمي الأول</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            
                            <p><b>الموضوع: الإنذار الأكاديمي الأول  بسبب عدم حضور اللقاءات المباشرة</b></p>
                            
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            
                            <p>أود أن أحيطكم علماً بأنه بسبب غيابكم عن حضور لقاءين  من اللقاءات المباشرة في مادة (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") التي أقوم بتدريسها، ولعدم تقديمكم عذراً مقبولاً لذلك، فإني أوجه لكم الإنذار الأكاديمي الأول  بناءً على ما ورد 
                            في القاعدة التنفيذية للمادة الرابعة عشرة من نظام الدراسة والامتحانات،والتي تنص على أنه يوجه للطالب المتغيب عن لقاءين رسالة إنذار أول.
                    كما أحيطكم علماً بأن القاعدة التنفيذية المشار إليها تنص كذلك على أنه يوجه للطالب المتغيب عن لقاءين رسالة إنذار ثانٍ.
                    علماً بأنه إذا قلت نسبة حضورك عن نسبة (70%) من مجموع حضور المحاضرات العامة والمحاضرات التطبيقية المحددة وأداء نشاطات التقييم المستمر في التعليم عن بعد؛ لكل مادة دراسية خلال الفصل الدراسي، فسيوجّه إليك رسالة حرمان من 
                            دخول الامتحان النهائي.</p>
                            
                            <p>آملين منكم حسن متابعة دراستكم، والالتزام بالأنظمة الأكاديمية </p>
                            
                            <p>داعين الله تعالى لك بالتوفيق والنجاح.</p>
                            
                            <br /><br /><br />
                            
                            محاضر المادة<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    SaveEmailMessage(
                        subjectRegistered.Student, subjectRegistered.TutorialGroup.Tutor.Profile.Email,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "FIRST ATTENDANCE WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        public bool SendBachelorAttendanceFirstWarningLetterOnCampus(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = subjectRegistered.TutorialGroup.Tutor.Profile.Email;
            }

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "First Academic Warning (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""http://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>THE FIRST ACADEMIC WARNING</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            
                            <p><b>Subject: The First Academic Warning due to the Non-attendance of Direct Meetings</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>Please be informed that due to your failure to attend the direct meetings of the subject (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") that 
                            I am lecturing for two times, and due to the fact that there is no legitimate excuse submitted by you, I, hereby, 
                            give you the first academic warning coinciding with the Article no.14 of the Studying and Examination 
                            Policy, which stated that the student who failed to attend the direct meetings twice must be given the 
                            First Academic warning.</p>
                            
                            <p>I am also bringing to your attention, that the above Article defines that the student whose absence 
                            reached three times he will receive the Second Academic Warning.</p>
                            
                            <p>For your information, if the percentage of your attendance is less than 80% from the total attendance 
                            of the total practical lectures and the overall participation in all Evaluated 
                            Activities of On-Campus Teaching in every subject each semester, the letter of Prohibition from Attending the 
                            Final Exam will be issued to you.</p>
                            
                            <p>We hope that you will adhere to your study attendance, and obey the academic policies.</p>
                            
                            <p>We pray to Allah Almighty to give you all the success and prosperity.</p>
                            
                            <br /><br /><br />
                            Lecturer of the subject<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>الإنذار الأكاديمي الأول</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            
                            <p><b>الموضوع: الإنذار الأكاديمي الأول  بسبب عدم حضور اللقاءات المباشرة</b></p>
                            
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            
                            <p>أود أن أحيطكم علماً بأنه بسبب غيابكم عن حضور لقائين  في مادة  (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") التي أقوم بتدريسها، ولعدم تقديمكم عذراً مقبولاً لذلك، فإني أوجه لكم الإنذار الأكاديمي الأول بناءً على ما ورد  
                            في القاعدة التنفيذية للمادة الرابعة عشرة من نظام الدراسة والامتحانات،والتي تنص على أنه يوجه للطالب المتغيب عن لقاءين رسالة إنذار أول.
                    كما أحيطكم علماً بأن القاعدة التنفيذية المشار إليها تنص كذلك على أنه يوجه للطالب المتغيب عن ثلاثة لقاءات رسالة إنذار ثانٍ.
                    علماً بأنه إذا قلت نسبة حضورك عن نسبة (80%) من مجموع حضور المحاضرات العامة والمحاضرات التطبيقية المحددة وأداء نشاطات التقييم المستمر في التعليم المباشر  لكل مادة دراسية خلال الفصل الدراسي، فسيوجّه إليك رسالة حرمان من 
                             دخول الامتحان النهائي.</p>
                            
                            <p>آملين منكم حسن متابعة دراستكم، والالتزام بالأنظمة الأكاديمية </p>
                            
                            <p>داعين الله تعالى لك بالتوفيق والنجاح.</p>
                            
                            <br /><br /><br />
                            
                            محاضر المادة<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    SaveEmailMessage(
                        subjectRegistered.Student, subjectRegistered.TutorialGroup.Tutor.Profile.Email,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "FIRST ATTENDANCE WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        public bool SendBachelorAttendanceSecondWarningLetter(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = subjectRegistered.TutorialGroup.Tutor.Profile.Email;
            }

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "Second Academic Warning (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>THE SECOND ACADEMIC WARNING</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            
                            <p><b>Subject: Second Academic Warning due to the Non-attendance of Direct Meetings</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>Please be informed, that due to your failure to attend the direct meetings of 
                            the subject (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") that I am lecturing for three times, and due to the fact that there is no legitimate excuse submitted by 
                            you, I, hereby,    give you the second academic warning coinciding with the Article no 14 of the Studying and 
                            Examination Policy, which stated that the student who failed to attend the direct meetings thrice must be 
                            given the Second Academic warning.</p>
                            
                            <p>I am also bringing to your attention, that the above article defines that the student whose absence reached 
                            three times he will receive the Second Academic Warning.</p>
                            
                            <p>For your information, if the percentage of your attendance is less than 70% from the total attendance of the 
                            total practical lectures and the overall participation in all Evaluated Activities of 
                            On-Line Teaching in every subject each semester, the letter of Prohibition from Attending the Final Exam will be 
                            issued to you.</p>
                            
                            <p>We hope that you will adhere to your study attendance, and obey the academic policies.</p>
                            
                            <p>We pray to Allah Almighty to give you all the success and prosperity.</p>
                            
                            <br /><br /><br />
                            Lecturer of the subject<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>الإنذار الأكاديمي الثاني</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            
                            <p><b>الموضوع: الإنذار الأكاديمي  الثاني بسبب عدم حضور اللقاءات المباشرة</b></p>
                            
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            
                            <p>أود أن أحيطكم علماً بأنه بسبب غيابكم عن حضور ثلاث لقاءات من اللقاءات المباشرة في مادة (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") التي أقوم بتدريسها، ولعدم تقديمكم عذراً مقبولاً لذلك، فإني أوجه لكم إنذاراً أكاديمياً ثانيا بناءً على ما ورد في القاعدة التنفيذية للمادة الرابعة عشرة من نظام الدراسة والامتحانات، والتي تنص على أنه يوجه للطالب المتغيب عن لقاءين رسالة إنذار ثان.
                    كما أحيطكم علماً بأن القاعدة التنفيذية المشار إليها تنص كذلك على أنه  يوجه للطالب المتغيب عن أكثر من ثلاثة لقاءات 
                    أو قلت نسبة حضوره عن نسبة (70%) من مجموع حضور المحاضرات العامة والمحاضرات التطبيقية المحددة وأداء نشاطات التقييم المستمر في التعليم عن بعد؛ لكل مادة دراسية خلال الفصل الدراسي، فسيوجّه إليه رسالة حرمان من دخول الامتحان النهائي.
                    </p>
                            
                            <p>آملين منكم حسن متابعة دراستكم، والالتزام بالأنظمة الأكاديمية </p>
                            
                            <p>داعين الله تعالى لك بالتوفيق والنجاح.</p>
                            
                            <br /><br /><br />
                            
                            محاضر المادة<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    SaveEmailMessage(
                        subjectRegistered.Student, subjectRegistered.TutorialGroup.Tutor.Profile.Email,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "SECOND ATTENDANCE WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        public bool SendBachelorAttendanceSecondWarningLetterOnCampus(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = subjectRegistered.TutorialGroup.Tutor.Profile.Email;
            }

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "Second Academic Warning (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>THE SECOND ACADEMIC WARNING</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            
                            <p><b>Subject: Second Academic Warning due to the Non-attendance of Direct Meetings</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>Please be informed, that due to your failure to attend the direct meetings of 
                            the subject (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") that I am lecturing for three times, and due to the fact that there is no legitimate excuse submitted by 
                            you, I, hereby, give you the second academic warning coinciding with the Article no 14 of the Studying and 
                            Examination Policy, which stated that the student who failed to attend the direct meetings thrice must be 
                            given the Second Academic Warning.</p>
                            
                            <p>I am also bringing to your attention, that the above article defines that the student whose absence reached 
                            three times he will receive the Second Academic Warning.</p>
                            
                            <p>For your information, if the percentage of your attendance is less than 80% from the total attendance 
                            of the total practical lectures and the overall participation in all Evaluated 
                            Activities of On-Campus Teaching in every subject each semester, the letter of Prohibition from Attending the 
                            Final Exam will be issued to you.</p>
                            
                            <p>We hope that you will adhere to your study attendance, and obey the academic policies.</p>
                            
                            <p>We pray to Allah Almighty to give you all the success and prosperity.</p>
                            
                            <br /><br /><br />
                            Lecturer of the subject<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>الإنذار الأكاديمي الثاني</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            
                            <p><b>الموضوع: الإنذار الأكاديمي  الثاني بسبب عدم حضور اللقاءات المباشرة</b></p>
                            
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            
                            <p>أود أن أحيطكم علماً بأنه بسبب غيابكم عن حضور ثلاثة لقاءات في مادة  (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") التي أقوم بتدريسها، ولعدم تقديمكم عذراً مقبولاً لذلك، فإني أوجه لكم الإنذار الأكاديمي الثاني  بناءً على ما ورد في القاعدة التنفيذية للمادة الرابعة عشرة من نظام الدراسة والامتحانات،والتي تنص على أنه يوجه للطالب المتغيب عن ثلاثة لقاءات رسالة إنذار ثان.
                    علماً بأنه إذا قلت نسبة حضورك عن نسبة (80%) من مجموع حضور المحاضرات العامة والمحاضرات التطبيقية المحددة وأداء نشاطات التقييم المستمر في التعليم المباشر لكل مادة دراسية خلال الفصل الدراسي، فسيوجّه إليك رسالة حرمان من دخول الامتحان النهائي.
                    </p>
                            
                            <p>آملين منكم حسن متابعة دراستكم، والالتزام بالأنظمة الأكاديمية </p>
                            
                            <p>داعين الله تعالى لك بالتوفيق والنجاح.</p>
                            
                            <br /><br /><br />
                            
                            محاضر المادة<br />
                            (" + subjectRegistered.TutorialGroup.Tutor.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    SaveEmailMessage(
                        subjectRegistered.Student, subjectRegistered.TutorialGroup.Tutor.Profile.Email,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "SECOND ATTENDANCE WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        public bool SendBachelorAttendanceProhibitLetter(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            if (subjectRegistered.TutorialGroup.Tutor.Department == null)
            {
                throw new Exception("The tutor has no department.");
            }
            var dean = subjectRegistered.TutorialGroup.Tutor.Department.Faculty.Deanship;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = dean.Profile.Email;
            }

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "Prohibition Attending Final Exam (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>PROHIBITION FROM ATTENDING THE FINAL EXAM</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            
                            <p><b>Subject: Notification of Prohibition from Attending the Final Exam</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>I would like to draw your attention, that the percentage of your overall attendance of the total
                            practical lectures and the overall participation in all Evaluated Activities of On-Line Teaching 
                            in every subject each semester is less than 70%. Hence in coinciding with the Article No: 14th of the Policy 
                            of Studying and Examination it is decided to prohibit you from attending the final exam of 
                            the subject (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @").</p>
                            
                            <br /><br /><br />
                            Dean of the faculty<br />
                            (" + dean.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>إشعار الحرمان من دخول الامتحان النهائي</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            
                            <p><b>الموضوع: إشعار بالحرمان من دخول الامتحان النهائي</b></p>
                            
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            
                            <p>أود أن أحيطكم علماً بأنه بسبب أن نسبة حضوركم قلت عن نسبة (70%) من مجموع حضور المحاضرات العامة والمحاضرات التطبيقية المحددة وأداء نشاطات التقييم المستمر في التعليم عن بعد؛ فعليه سيتم حرمانكم من دخول الامتحان النهائي.</p>
                            
                            <p>في مادة (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") بناءً على ما ورد في القاعدة التنفيذية للمادة الرابعة عشرة من نظام الدراسة والامتحانات.</p>
                            
                            <br /><br /><br />
                            
                            عميد الكلية<br />
                            (" + dean.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    string senderEmail = "";
                    if (dean != null)
                        senderEmail = dean.Profile.Email;
                    SaveEmailMessage(
                        subjectRegistered.Student, senderEmail,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "PROHIBIT ATTENDANCE WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        public bool SendBachelorAttendanceProhibitLetterOnCampus(SubjectRegistered subjectRegistered, string emailTo)
        {
            string recipientEmail;
            if (subjectRegistered.TutorialGroup.Tutor.Department == null)
            {
                throw new Exception("The tutor has no department.");
            }
            var dean = subjectRegistered.TutorialGroup.Tutor.Department.Faculty.Deanship;
            if (emailTo == "student")
            {
                recipientEmail = subjectRegistered.Student.Profile.Email;
            }
            else
            {
                recipientEmail = dean.Profile.Email;
            }

            var faculty = GetFaculty(subjectRegistered.Student.Course.CourseDescription.Faculty);
            var subject = "Prohibition Attending Final Exam (due to the Non-attendance of Direct Meetings)";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                        </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">
                            <center><h3>PROHIBITION FROM ATTENDING THE FINAL EXAM</h3></center>
                            <hr />
                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + subjectRegistered.Student.Profile.Name + @"<br />
                            Matriculation No: " + subjectRegistered.Student.MatrixNo + @"<br />
                            Faculty: " + faculty.DisplayName + @"<br />
                            
                            <p><b>Subject: Notification of Prohibition from Attending the Final Exam</b></p>
                            
                            <p>Peace, Mercy and ALLAH’ Blessing Upon you;</p>
                            
                            <p>I would like to draw your attention, that the percentage of your overall attendance of the total 
                            practical lectures, and the overall participation in all Evaluated Activities of On-Campus Teaching 
                            is less than 80%. Hence in coinciding with the Article No: 14th of the Policy 
                            of Studying and Examination it is decided to prohibit you from attending the final exam of 
                            the subject (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @").</p>
                            
                            <br /><br /><br />
                            Dean of the faculty <br />
                            (" + dean.Profile.Name + @")        
                            </td>
                            
                            <td width=""50%"" dir=""rtl"" valign=""top"">
                            <center><h3>إشعار الحرمان من دخول الامتحان النهائي</h3></center>
                            <hr />
                            التاريخ: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            اسم الطالب / الطالبة: " + subjectRegistered.Student.Profile.Name + @"<br />
                            الرقم المرجعي: " + subjectRegistered.Student.MatrixNo + @"<br />
                            الكلية: " + faculty.DisplayName + @"<br />
                            
                            <p><b>الموضوع: إشعار بالحرمان من دخول الامتحان النهائي</b></p>
                            
                            <p>السلام عليكم ورحمة الله وبركاته،             وبعد:</p>
                            
                            <p>أود أن أحيطكم علماً بأنه بسبب أن نسبة حضوركم قلت عن (80%) من مجموع حضور المحاضرات العامة والمحاضرات التطبيقية المحددة وأداء نشاطات التقييم المستمر في التعليم المباشر ؛ فعليه سيتم حرمانكم من دخول الامتحان النهائي                             
                            في مادة (" + subjectRegistered.Subject.Code + " - " + subjectRegistered.Subject.NameEn + @") بناءً على ما ورد في القاعدة التنفيذية للمادة الرابعة عشرة من نظام الدراسة والامتحانات.</p>
                            
                            <br /><br /><br />
                            
                            عميد الكلية<br />
                            (" + dean.Profile.Name + @")
                            
                            </td>
                        </tr>
                    </table>
                       ";
            //return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            if (svcMessaging.SendEmail(recipientEmail, subject, body, true))
            {
                if (emailTo == "student")
                {
                    string senderEmail = "";
                    if (dean != null)
                        senderEmail = dean.Profile.Email;
                    SaveEmailMessage(
                        subjectRegistered.Student, senderEmail,
                        DescriptionForStudentAcademicEmail(subjectRegistered),
                        body, "PROHIBIT ATTENDANCE WARNING", subject, "attendance.html");
                }
                return true;
            }
            return false;
        }

        // Second Version Scholarship Email
        public bool SendFullScholarshipOfferLetter(ScholarshipRegistered sr, string acceptanceBaseUrl, Profile personTriggerProfile)
        {
            acceptanceBaseUrl += "?u=" + sr.ScholarshipApply.Student.Profile.Username + "&i=" + sr.Id;
            string subject = "Full Scholarship Grant Offer";
            var body = @"
                        <div><center><img src=""http://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div>
                        <br />
                        <table width=""100%"">
                            <tr>
                                <td width=""50%"" dir=""ltr"">
                                    <p>
                                        <b>Date : </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                        <b>Name : </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                        <b>Matric No : </b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                        <b>Reference No : </b> " + sr.ScholarshipApply.RefNo + @" <br />
                                    </p><br />
                                    <p>
                                        <b>Subject: Full Scholarship Grant Offer</b>
                                    </p><br />
                                    
                                    <p>Al-Madinah International University is pleased to inform you that as a student you have been 
                                    granted a full scholarship for this current semester February 2013. Accordingly, if you agree to 
                                    this, please notify us officially by clicking on the icon below “I Accept”. Please note that any delay 
                                    in giving such consent,  implies us you are refusing such scholarship offer.</p>
                                </td>
                                <td width=""50%"" dir=""rtl"">
                                    <p>
                                        <b>تاريخ الرسالة: </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                        <b>اسم الطالب: </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                        <b>رقم الطالب المرجعي :</b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                        <b>رقم المنحة المرجعي :</b> " + sr.ScholarshipApply.RefNo + @" <br />
                                    </p><br />
                                    <p>
                                        <b>الموضوع: عرض كفالة منحة دراسية كاملة</b>
                                    </p><br />
                                    
                                    <p>يسر جامعة المدينة العالمية أن تشعركم بموافقتها على إدراجكم ضمن الطلبة الحاصلين على منحة دراسية كاملة للفصل الدراسي الحالي  فبراير 2013.
                    وعليه فيرجى اشعارنا رسميا فى حالة موافقتكم عن طريق الضغط على أيقونة أقبل بالعرض أدناه..  علما بأن التأخر في إبداء هذه الموافقة يعني ضمنيا رفضكم لعرض كفالة المنحة الدراسية.
                                    </p>
                                </td>
                            </tr>
                        </table>
                        
                        <center><p><a href=""" + acceptanceBaseUrl + @""">أقبل بالعرض I Accept / </a></p></center>
                        
                        <table width=""100%"">
                            <tr>
                                <td width=""50%"" dir=""ltr"">
                                    <p>Since you were chosen from among a large number of applicants. We hope that this scholarship enable
                                     you boost up your morale and encourage you to work hardly during your studies at the university. 
                                     May Allah bless your success.
                                    </p><br />
                                    
                                    <p>Grants Committee Unit<br />
                                    " + personTriggerProfile.Name + @"<br />
                                    Al-Madinah International University
                                    </p>
                                </td>
                                  <td width=""50%"" dir=""rtl"">
                                  <p>وحيث أنه تم اختياركم من بين عدد كبير من المتقدمين.. فإننا نأمل أن تعينكم هذه المنحة في رفع الروح المعنوية لديكم ، وحثكم على الجد والاجتهاد والانضباط أثناء دراستكم بالجامعة. والله ولي التوفيق.</p><br />
                                  
                                  <p>لجنة وحدة المنح<br />
                                  " + personTriggerProfile.Name + @"<br />
                                  جامعة المدينة العالمبة
                                  </p>
                                </td>
                            </tr>
                        </table>
                        ";
            //sr.ScholarshipApply.Student.Profile.Email
            return svcMessaging.SendEmail(sr.ScholarshipApply.Student.Profile.Email, subject, body, true);
        }

        public bool SendPartialScholarshipOfferLetter(ScholarshipRegistered sr, string acceptanceBaseUrl, string refuseBaseUrl, Profile personTriggerProfile)
        {
            acceptanceBaseUrl += "?u=" + sr.ScholarshipApply.Student.Profile.Username + "&i=" + sr.Id;
            refuseBaseUrl += "?u=" + sr.ScholarshipApply.Student.Profile.Username + "&i=" + sr.Id;
            string subject = "Tuition Fees Grant Offer";
            string amountToPayInMYR = "MYR " + sr.ScholarshipFee.FinalStudentNeedToPay("MYR").ToString("0.00");
            string amountToPayInUSD = "USD " + sr.ScholarshipFee.FinalStudentNeedToPay("USD").ToString("0.00");
            string scholarshipAmountMYR = "MYR " + sr.ScholarshipFee.FinalGrandTotal("MYR").ToString("0.00");
            string scholarshipAmountUSD = "USD " + sr.ScholarshipFee.FinalGrandTotal("USD").ToString("0.00");

            string baseUrl = "http://cms.mediu.edu.my/office/";
            //string baseUrl = "http://localhost:55555/";
            var body = @"
                            <div><center><img src='" + baseUrl + @"content/MIF.png' alt=""mediu"" /></center></div>
                            <br />
                            <table width=""100%"" dir=""ltr"">
                                <tr>
                                    <td width=""50%"" dir=""ltr"">
                                        <p>
                                            <b>Date : </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>Name : </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>Matric No : </b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>Reference No : </b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>Subject: Offer of Conditional Partial Scholarship</b>
                                        </p><br />
                                        
                                        <p>Greetings from Al-Madinah International Foundation (MIF)</p>
                                        <p>Within the framework of existing cooperation between Al-Madinah International Foundation (MIF) and the Al-Madinah International University (MEDIU) with regard to sponsoring some students to achieve their dreams and aspirations to complete their tertiary education in higher education institutions, especially those who has proven their outstanding academic performance but, at the same time lack of sufficient funding to further their tertiary education at Al-Madinah International University (MEDIU).

                                        Al-Madinah International Foundation (MIF) has sincere willingness to help those students in dire needs, the Foundation established a scholarship program through which Al-Madinah International Foundation (MIF) acts as a mediator between you (as a student) and the donors who want to sponsor Students with outstanding academic grades/performance.

                                        Based on the request for financial aid in the form of scholarship submitted by yourself as a student. Al-Madinah International Foundation MIF have to look into these scholarship application to be submitted to Al-Madinah International University (MEDIU), and MIF has decided to offer partial scholarship with the following conditions:
                                        </p>
                                        <p>Firstly: the partial scholarship to be offered of " + scholarshipAmountMYR + @", from the value of the tuition fees for the current one semester only.
                                        </p>
                                        <p>Secondly: this offer (of Partial Scholarship) is subject to pay all your other financial obligations to  Al-Madinah International University (MEDIU) as per latest amount dues (arrears) in your financial invoice of " + amountToPayInMYR + @", issued to you in this current semester.
                                        </p>
                                        <p>Thirdly: The Partial Scholarship Grant is for one semester only, and the renewal is subject to financial standing of MIF, and your commitment to continue your studies with regard to your attendance, submission of assignment, follow-up activities, and other duties as a student as well as commitment to ensure maintenance of your high CGPA.
                                        </p>
                                        <p>Fourthly: the priority in partial scholarship will be for students who have come forward to pay all payment dues to the University (admin fees and other related fees), in order to help this institution to cover its operational expenses.
                                        </p>
                                        <p>Fifthly: This offer is valid for a maximum of one month from the date of this discourse.
                                        </p>
                                        <br />
                                        <p>Director for Al-Madinah International Foundation (MIF)</p>
                                        <p>Hj Bahrudin A. Rahman</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>
                                    </td>
                                    <td style='font-size:15px;' width=""50%"" dir=""rtl"">
                                        <p>
                                            <b>تاريخ الرسالة: </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>اسم الطالب: </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>رقم الطالب المرجعي :</b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>رقم المنحة المرجعي :</b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>الموضوع: عرض منحة مشروط</b>
                                        </p><br />
                                        
                                        <p>تحية طيبة من مؤسسة المدينة العالمية الخيرية ( ماليزيا )</p>
                                        <br/><p>في إطار التعاون القائم بين مؤسسة المدينة العالمية الخيرية وجامعة المدينة العالمية  بخصوص كفالة بعض الطلاب لتحقيق أحلامهم وطموحاتهم بإكمال تعليمهم في مؤسسات التعليم العالي، وبخاصة أولئك الذين يتمتعون بالأداء الأكاديمي المتميز وفى الوقت نفسه يفتقرون إلي التمويل الكافي لمواصلة تعليمهم الجامعي في جامعة المدينة العالمية.</p>
                                        <br/><p>ورغبة من مؤسسة المدينة العالمية الخيرية في مساعدة أولئك الطلاب فقد قامت المؤسسة بإنشاء برنامج للمنح الدراسية تقوم من خلاله المؤسسة بدور الوسيط بينكم وبين المانحين الذين يرغبون بكفالة الطلاب أصحاب المعدلات الدراسية المتميزة.</p>
                                        <br/><p>وبناءً على طلب المنحة المقدم منكم فقد قامت مؤسسة المدينة العالمية الخيرية بدراسة طلب المنحة المقدم من قبلكم إلي جامعة المدينة العالمية، وعليه تقرر تقديم عرض منحة جزئية مشروط بالأمور الآتية:</p>
                                        <br/><p>أولاً: أن يكون عرض المنحة الجزئية الخاص بكم   بقيمة " + scholarshipAmountMYR + @" وذلك من قيمة رسوم المواد الدراسية الخاصة بهذا الفصل فقط. </p>
                                        <p>ثانياً: أن هذا العرض مشروط بسدادكم لجميع الالتزامات المالية الخاصة بكم تجاه جامعة المدينة العالمية بحسب المبلغ المتبقي من الفاتورة الصادرة من الجامعة لهذا الفصل وقدره" + amountToPayInMYR + @".</p>
                                        <p>ثالثاً: أن عرض المنحة المقدم لكم هو لفصل دراسي واحد فقط، وتجديده خاضع لإمكانات المؤسسة المالية، ولالتزامكم بمواصلة الدراسة والحرص على حضور اللقاءات المباشرة ومتابعة الأنشطة والواجبات والتزامكم بالحصول على معدل تراكمي مرتفع.</p>
                                        <p>رابعاً: أن الأولوية في المنح ستكون للطلبة الذين يبادرون بسداد ما عليهم تجاه الجامعة، وذلك بغية مساعدة الجامعة في تغطية مصروفاتها التشغيلية.</p>
                                        <p>خامساً: أن هذا العرض صالح لمدة أقصاها شهر من تاريخ خطابنا هذا. </p>
                                        <br />
                                        <p>مدير عام مؤسسة المدينة العالمية الخيرية</p>
                                        <p>الحاج بحر الدين عبد الرحمن</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>                                    
                                        
                                    </td>
                                </tr>
                            </table><br />
                            
                            <center><p><a href=""" + acceptanceBaseUrl + @""">Agree, I shall pay admin fees and other related fees due to the University / أقبل بالعرض ، وسأقوم بسداد باقي الرسوم المتبقية للجامعة </a></p></center>
                            <center><p><a href=""" + refuseBaseUrl + @""">I disagree, the offer of this partial scholarship is to be cancelled automatically / لا أوافق، مع علمي بأن طلب المنحة سيتم إلغاؤه </a></p></center>
                           
                        ";
            //sr.ScholarshipApply.Student.Profile.Email
            return svcMessaging.SendEmail(sr.ScholarshipApply.Student.Profile.Email, subject, body, true);
        }

        public bool SendRejectionScholarshipEmail(ScholarshipRegistered sr, Profile personTriggerProfile, string choice)
        {
            string subject = "Scholarship Rejection Letter";
            string baseUrl = "http://cms.mediu.edu.my/office/";
            string armessage = "";
            string enmessage = "";

            //cgpaMa' /> You got a cumulative average less than “3.00” - Masters Level<br />
            //cgpaBa' /> You got a cumulative average less than “2.50” Bachelor level<br />
            //parttime' /> The scholarship shall not be granted to a student who is studying part-time.<br />
            //onlyBaAndMa' /> The scholarship is for the students of undergraduate studies and postgraduate studies only.<br />
            switch (choice)
            {
                case "cgpaMa":
                    armessage = "حصولكم على معدل تراكمي يقل عن 3.00 في مرحلة الماجستير";
                    enmessage = "You got a cumulative average less than “3.00” - Masters Level";
                    break;
                case "cgpaBa":
                    armessage = "حصولكم على معدل تراكمي يقل عن 2.50 في مرحلة البكالوريوس";
                    enmessage = "You got a cumulative average less than “2.50” Bachelor level";
                    break;
                case "parttime":
                    armessage = "أن المنحة لا تمنح للطالب الذي يدرس بنظام الدوام الجزئي";
                    enmessage = "The scholarship shall not be granted to a student who is studying part-time";
                    break;
                case "onlyBaAndMa":
                    armessage = "أن المنح حاليا لطلاب الدراسات الجامعية والدراسات العليا فقط";
                    enmessage = "The scholarship is for the students of undergraduate studies and postgraduate studies only";
                    break;

            }
            //string baseUrl = "http://localhost:55555/";
            var body = @"
                            <div><center><img src='" + baseUrl + @"content/MIF.png' alt=""mediu"" /></center></div>
                            <br />
                            <table width=""100%"" dir=""ltr"">
                                <tr>
                                    <td width=""50%"" dir=""ltr"">
                                        <p>
                                            <b>Date : </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>Name : </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>Matric No : </b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>Reference No : </b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>Subject: Notice of refusing the scholarship</b>
                                        </p><br />
                                        
                                        <p>Greetings from Al-Madinah International Charitable Foundation</p>
                                        <p>Based on the scholarship application submitted by you, Al-Madinah International Charitable Foundation (MIF) studied your scholarship application submitted to Al-Madinah International University (MEDIU), Due to donor’s decision to give a scholarship to only students under certain conditions it was decided to refuse your request for the scholarship because of:</p>
                                        <br />
                                        <p>
                                        " + enmessage + @"
                                        </p>
                                        <br />
                                        <p>With our best wishes in successes and reconcile </p>
                                        <p>Director for Al-Madinah International Foundation (MIF)</p>
                                        <p>Hj Bahrudin A. Rahman</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>
                                    </td>
                                    <td style='font-size:15px;' width=""50%"" dir=""rtl"">
                                        <p>
                                            <b>تاريخ الرسالة: </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>اسم الطالب: </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>رقم الطالب المرجعي :</b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>رقم المنحة المرجعي :</b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>الموضوع: إشعار برفض المنحة الدراسية</b>
                                        </p><br />
                                        
                                        <p>تحية طيبة من مؤسسة المدينة العالمية الخيرية ( ماليزيا )</p>
                                        <br/><p>بناءً على طلب المنحة المقدم منكم فقد قامت مؤسسة المدينة العالمية الخيرية(MIF)  بدراسة طلب المنحة المقدم من قبلكم إلي جامعة المدينة العالمية (MEDIU) ، ونظراً لرغبة المانحين بكفالة الطلاب بشروط محددة فقد تقرر رفض طلب المنحة الخاص بكم وذلك بسبب: </p>                                        
                                        <br/>
                                        <p>
                                        " + armessage + @"
                                        </p>
                                        <br /><p>مع أطيب تمنياتنا لكم بدوام النجاح والتوفيق</p>
                                        <p>مدير عام مؤسسة المدينة العالمية الخيرية</p>
                                        <p>الحاج بحر الدين عبد الرحمن</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>                                    
                                        
                                    </td>
                                </tr>
                            </table><br />
                        ";
            //sr.ScholarshipApply.Student.Profile.Email
            return svcMessaging.SendEmail(sr.ScholarshipApply.Student.Profile.Email, subject, body, true);
        }

        public bool SendSecondRejectionScholarshipEmail(ScholarshipRegistered sr, Profile personTriggerProfile, string choice)
        {
            string subject = "Scholarship Rejection Letter";
            string baseUrl = "http://cms.mediu.edu.my/office/";
            string armessage = "";
            string enmessage = "";

            //cgpaMa' /> You got a cumulative average less than “3.00” - Masters Level<br />
            //cgpaBa' /> You got a cumulative average less than “2.50” Bachelor level<br />
            //parttime' /> The scholarship shall not be granted to a student who is studying part-time.<br />
            //onlyBaAndMa' /> The scholarship is for the students of undergraduate studies and postgraduate studies only.<br />
            switch (choice)
            {
                case "attendance":
                    armessage = "عدم الانتظام في الدراسة بشكل مرض أو لضعف أداء الأنشطة";
                    enmessage = "Irregular attendance or unsatisfactory activities performance";
                    break;
                case "withdrawal":
                    armessage = "تكرر تأجيل الدراسة أو الانسحاب أو التعليق في المواسم الماضية";
                    enmessage = "Repeated deferment, withdrawal, or suspension of study in previous semesters";
                    break;
                case "termination":
                    armessage = "الفصل أو التأجيل أو التعليق لقيدكم الدراسي بعد التقديم لطلب المنحة";
                    enmessage = "Termination, deferment, or suspension of enrollment after application for a scholarship";
                    break;
            }
            //string baseUrl = "http://localhost:55555/";
            var body = @"
                            <div><center><img src='" + baseUrl + @"content/MIF.png' alt=""mediu"" /></center></div>
                            <br />
                            <table width=""100%"" dir=""ltr"">
                                <tr>
                                    <td width=""50%"" dir=""ltr"">
                                        <p>
                                            <b>Date : </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>Name : </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>Matric No : </b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>Reference No : </b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>Subject: Notice of refusing the scholarship</b>
                                        </p><br />
                                        
                                        <p>Greetings from Al-Madinah International Charitable Foundation</p>
                                        <p>Based on the scholarship application submitted by you, Al-Madinah International Charitable Foundation (MIF) studied your scholarship application submitted to Al-Madinah International University (MEDIU), Due to donor’s decision to give a scholarship to only students under certain conditions it was decided to refuse your request for the scholarship because of:</p>
                                        <br />
                                        <p>
                                        " + enmessage + @"
                                        </p>
                                        <br />
                                        <p>With our best wishes in successes and reconcile </p>
                                        <p>Director for Al-Madinah International Foundation (MIF)</p>
                                        <p>Hj Bahrudin A. Rahman</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>
                                    </td>
                                    <td style='font-size:15px;' width=""50%"" dir=""rtl"">
                                        <p>
                                            <b>تاريخ الرسالة: </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>اسم الطالب: </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>رقم الطالب المرجعي :</b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>رقم المنحة المرجعي :</b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>الموضوع: إشعار برفض المنحة الدراسية</b>
                                        </p><br />
                                        
                                        <p>تحية طيبة من مؤسسة المدينة العالمية الخيرية ( ماليزيا )</p>
                                        <br/><p>بناءً على طلب المنحة المقدم منكم فقد قامت مؤسسة المدينة العالمية الخيرية(MIF)  بدراسة طلب المنحة المقدم من قبلكم إلي جامعة المدينة العالمية (MEDIU) ، ونظراً لرغبة المانحين بكفالة الطلاب بشروط محددة فقد تقرر رفض طلب المنحة الخاص بكم وذلك بسبب: </p>                                        
                                        <br/>
                                        <p>
                                        " + armessage + @"
                                        </p>
                                        <br /><p>مع أطيب تمنياتنا لكم بدوام النجاح والتوفيق</p>
                                        <p>مدير عام مؤسسة المدينة العالمية الخيرية</p>
                                        <p>الحاج بحر الدين عبد الرحمن</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>                                    
                                        
                                    </td>
                                </tr>
                            </table><br />
                        ";
            //sr.ScholarshipApply.Student.Profile.Email
            return svcMessaging.SendEmail(sr.ScholarshipApply.Student.Profile.Email, subject, body, true);
        }

        public bool SendThirdRejectionScholarshipEmail(ScholarshipRegistered sr, Profile personTriggerProfile, string choice)
        {
            string subject = "Scholarship Rejection Letter";
            string baseUrl = "http://cms.mediu.edu.my/office/";
            string armessage = "";
            string enmessage = "";

            //cgpaMa' /> You got a cumulative average less than “3.00” - Masters Level<br />
            //cgpaBa' /> You got a cumulative average less than “2.50” Bachelor level<br />
            //parttime' /> The scholarship shall not be granted to a student who is studying part-time.<br />
            //onlyBaAndMa' /> The scholarship is for the students of undergraduate studies and postgraduate studies only.<br />
            switch (choice)
            {
                case "notpaid":
                    armessage = "إلغاء العرض لعدم تسديد الجزء المتبقي من الفاتورة في الفترة المحددة.";
                    enmessage = "You have not paid the remaining invoice fees in the specified period";
                    break;
            }
            //string baseUrl = "http://localhost:55555/";
            var body = @"
                            <div><center><img src='" + baseUrl + @"content/MIF.png' alt=""mediu"" /></center></div>
                            <br />
                            <table width=""100%"" dir=""ltr"">
                                <tr>
                                    <td width=""50%"" dir=""ltr"">
                                        <p>
                                            <b>Date : </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>Name : </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>Matric No : </b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>Reference No : </b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>Subject: Notice of refusing the scholarship</b>
                                        </p><br />
                                        
                                        <p>Greetings from Al-Madinah International Charitable Foundation</p>
                                        <p>Based on the scholarship application submitted by you, Al-Madinah International Charitable Foundation (MIF) studied your scholarship application submitted to Al-Madinah International University (MEDIU), Due to donor’s decision to give a scholarship to only students under certain conditions it was decided to refuse your request for the scholarship because of:</p>
                                        <br />
                                        <p>
                                        " + enmessage + @"
                                        </p>
                                        <br />
                                        <p>With our best wishes in successes and reconcile </p>
                                        <p>Director for Al-Madinah International Foundation (MIF)</p>
                                        <p>Hj Bahrudin A. Rahman</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>
                                    </td>
                                    <td style='font-size:15px;' width=""50%"" dir=""rtl"">
                                        <p>
                                            <b>تاريخ الرسالة: </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>اسم الطالب: </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>رقم الطالب المرجعي :</b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>رقم المنحة المرجعي :</b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>الموضوع: إشعار برفض المنحة الدراسية</b>
                                        </p><br />
                                        
                                        <p>تحية طيبة من مؤسسة المدينة العالمية الخيرية ( ماليزيا )</p>
                                        <br/><p>بناءً على طلب المنحة المقدم منكم فقد قامت مؤسسة المدينة العالمية الخيرية(MIF)  بدراسة طلب المنحة المقدم من قبلكم إلي جامعة المدينة العالمية (MEDIU) ، ونظراً لرغبة المانحين بكفالة الطلاب بشروط محددة فقد تقرر رفض طلب المنحة الخاص بكم وذلك بسبب: </p>                                        
                                        <br/>
                                        <p>
                                        " + armessage + @"
                                        </p>
                                        <br /><p>مع أطيب تمنياتنا لكم بدوام النجاح والتوفيق</p>
                                        <p>مدير عام مؤسسة المدينة العالمية الخيرية</p>
                                        <p>الحاج بحر الدين عبد الرحمن</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>                                    
                                        
                                    </td>
                                </tr>
                            </table><br />
                        ";
            //sr.ScholarshipApply.Student.Profile.Email
            return svcMessaging.SendEmail(sr.ScholarshipApply.Student.Profile.Email, subject, body, true);
        }

        public bool SendFourthRejectionScholarshipEmail(ScholarshipRegistered sr, Profile personTriggerProfile, string choice)
        {
            string subject = "Scholarship Rejection Letter";
            string baseUrl = "http://cms.mediu.edu.my/office/";
            string armessage = "";
            string enmessage = "";

            //cgpaMa' /> You got a cumulative average less than “3.00” - Masters Level<br />
            //cgpaBa' /> You got a cumulative average less than “2.50” Bachelor level<br />
            //parttime' /> The scholarship shall not be granted to a student who is studying part-time.<br />
            //onlyBaAndMa' /> The scholarship is for the students of undergraduate studies and postgraduate studies only.<br />
            switch (choice)
            {
                case "notenough":
                    armessage = "نفاد المبلغ المرصود للمنح لهذا الموسم";
                    enmessage = "Allocated budget run out of funds for this semester";
                    break;
            }
            //string baseUrl = "http://localhost:55555/";
            var body = @"
                            <div><center><img src='" + baseUrl + @"content/MIF.png' alt=""mediu"" /></center></div>
                            <br />
                            <table width=""100%"" dir=""ltr"">
                                <tr>
                                    <td width=""50%"" dir=""ltr"">
                                        <p>
                                            <b>Date : </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>Name : </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>Matric No : </b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>Reference No : </b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>Subject: Notice of refusing the scholarship</b>
                                        </p><br />
                                        
                                        <p>Greetings from Al-Madinah International Charitable Foundation</p>
                                        <p>Based on the scholarship application submitted by you, Al-Madinah International Charitable Foundation (MIF) studied your scholarship application submitted to Al-Madinah International University (MEDIU), Due to donor’s decision to give a scholarship to only students under certain conditions it was decided to refuse your request for the scholarship because of:</p>
                                        <br />
                                        <p>
                                        " + enmessage + @"
                                        </p>
                                        <br />
                                        <p>With our best wishes in successes and reconcile </p>
                                        <p>Director for Al-Madinah International Foundation (MIF)</p>
                                        <p>Hj Bahrudin A. Rahman</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>
                                    </td>
                                    <td style='font-size:15px;' width=""50%"" dir=""rtl"">
                                        <p>
                                            <b>تاريخ الرسالة: </b> " + DateTime.Now.ToString("dd/MMMM/yyyy") + @" <br />
                                            <b>اسم الطالب: </b> " + sr.ScholarshipApply.Student.Profile.Name + @" <br />
                                            <b>رقم الطالب المرجعي :</b> " + sr.ScholarshipApply.Student.MatrixNo + @" <br />
                                            <b>رقم المنحة المرجعي :</b> " + sr.ScholarshipApply.RefNo + @" <br />
                                        </p><br />
                                        <p>
                                            <b>الموضوع: إشعار برفض المنحة الدراسية</b>
                                        </p><br />
                                        
                                        <p>تحية طيبة من مؤسسة المدينة العالمية الخيرية ( ماليزيا )</p>
                                        <br/><p>بناءً على طلب المنحة المقدم منكم فقد قامت مؤسسة المدينة العالمية الخيرية(MIF)  بدراسة طلب المنحة المقدم من قبلكم إلي جامعة المدينة العالمية (MEDIU) ، ونظراً لرغبة المانحين بكفالة الطلاب بشروط محددة فقد تقرر رفض طلب المنحة الخاص بكم وذلك بسبب: </p>                                        
                                        <br/>
                                        <p>
                                        " + armessage + @"
                                        </p>
                                        <br /><p>مع أطيب تمنياتنا لكم بدوام النجاح والتوفيق</p>
                                        <p>مدير عام مؤسسة المدينة العالمية الخيرية</p>
                                        <p>الحاج بحر الدين عبد الرحمن</p>
                                        <p><img src='" + baseUrl + @"content/bahrSignature.png' alt=""mediu"" /></p>                                    
                                        
                                    </td>
                                </tr>
                            </table><br />
                        ";
            //sr.ScholarshipApply.Student.Profile.Email
            return svcMessaging.SendEmail(sr.ScholarshipApply.Student.Profile.Email, subject, body, true);
        }

        public bool SendConfirmToPayAdminFeeToApplicant(AdmissionApply apply)
        {
            var subject = "Confirmation To Pay Admin Fee";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <b>Reference :</b>" + apply.RefNo + @"<br />
                                <b>Name :</b>" + apply.ApplicantName + @" <br />
                                <p>Dear Applicant, </p>
                                <p>السلام عليكم ورحمة الله وبركاته </p>

                                <p>Following your interest and commitment to study at our prestigious University beginning this " + apply.Intake.Code + @" intake, It was decided to allow you to register your subjects for the current intake, provided that you take pledge to pay all administrative and tuition fees required in case your application for either full or partial scholarship becomes unsuccessful. Knowing that the university keeps the right to suspend or cancel your studies in the absence of your payment under this pledge.</p>
                                <p>Please use the link below to confirm your position:</p>    
                                    <p>http://online.mediu.edu.my/apply/Applicant/ConfirmToPayAdminFee.aspx?lang=en</p>
                                
                                <p>We wish you best of luck and success</p>
                                <p>
                                Thank you
                                </p>
                                <p>
                                Admission & Registration<br>
                                Al-Madinah International University
                                </p>
                                    
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>أخي المتقدم/ أُختي المتقدمة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>حرصا من الجامعة على مصلحتكم في مواصلة الدراسة من بداية الفصل الدراسي؛ فقد تقرر تمكينكم من تسجيل المواد الدراسية لفصل " + apply.Intake.Code + @" شريطة أن تتعهدوا بدفع كافة الرسوم المطلوبة منكم في حالة عدم حصولكم على منحة دراسية كلية أو جزئية من قبل الجامعة، مع العلم أن للجامعة الاحتفاظ بحقها في تعليق قيدكم في حالة عدم التزامكم  بالسداد بموجب هذا التعهد.</p>
                            <p>وعليه يرجى اختيار أحد الخيارات الآتية</p>
                                <p>http://online.mediu.edu.my/apply/Applicant/ConfirmToPayAdminFee.aspx?lang=ar</p>
                            <p>مع تمنياتنا لكم بالتوفيق والنجاح</p>
                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            قسم القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";

            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true);
        }

        public void GenerateUndergraduateFinalOffer(string destOffer, AdmissionApply apply, string offerLetterFileName, DateTime? dateOfferLetterSent)
        {
            DateTime currentDateTime = (dateOfferLetterSent != null) ? (DateTime)dateOfferLetterSent : DateTime.Now;
            var s = sm.OpenSession();
            string sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Final Undergraduate.pdf";
            try
            {
                IntakeEvent semester = apply.Intake;
                var semesters = s.GetAll<CampusEvent>()
                                          .ThatHasChild(c => c.ParentEvent)
                                              .Where(c => c.Id).IsEqualTo(semester.ParentEvent.Id)
                                          .EndChild()
                                          .And(c => c.EventType).IsEqualTo("SME")
                                          .Execute().ToList();
                CampusEvent semesterDates = semesters.Where(c => c.Code.ToLower().StartsWith(semester.Code.Substring(0, 3).ToLower())).FirstOrDefault();

                AdmissionApply absApply = apply;

                int serial = semester.OfferLetterSerialNumber + 1;
                string cardnum = absApply.Applicant.Profile.IdCardNumber;
                if (cardnum == "nil")
                {
                    cardnum = "";
                }

                PdfReader r = new PdfReader(sourceOffer);
                string guid = Guid.NewGuid().ToString();
                PdfStamper stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate));
                PdfContentByte canvas = stamper.GetOverContent(1);
                PdfContentByte canvas2 = stamper.GetOverContent(2);
                PdfContentByte canvas3 = stamper.GetOverContent(3);
                BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                canvas.SetFontAndSize(bf, 8);
                canvas2.SetFontAndSize(bf, 8);

                Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                Font f4 = new Font(bf, 12, Font.ITALIC, BaseColor.BLACK);
                Font f5 = new Font(bf, 10, Font.ITALIC, BaseColor.BLACK);

                //English


                // 1 Serial Number
                string serialNumber = serial.ToString("0000000") + "(" + currentDateTime.Year.ToString().Substring(2, 2) + currentDateTime.Month.ToString("00") + ")";
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);

                // 2 Semester
                //semester.Month
                String semesterName = semester.MonthName.ToUpper() + " " + semester.Year.ToString();

                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_CENTER, new Phrase(semesterName, f5), 297, 729, 0);

                // 3 Reference Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.RefNo, f3), 200, 600, 0);

                // 4 Name
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.NameEnglish, f3), 200, 590, 0);


                // 5 Passport Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.IdCardNumber, f3), 200, 580, 0);


                // 6 Semester
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(semesterName, f3), 200, 570, 0);


                // 7 Faculty
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Course1.Faculty, f3), 200, 560, 0);

                // 8 Level
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Course1.CourseDescription.CourseLevel, f3), 200, 550, 0);

                // 9 Program
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Course1.NameEn, f3), 200, 540, 0);


                // 10 Workload




                // 11 Normal Period
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f3), 200, 530, 0);


                // 5 Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(currentDateTime.ToString("dd.MM.yyyy"), f3), 200, 520, 0);

                // 9 Classes Commence Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(semesterDates.StartDate.Value.ToString("dd MMMMMM yyyy"), f5), 90, 338, 0);

                // 8 Response Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_CENTER, new Phrase(currentDateTime.AddDays(14).ToString("dd MMMMMM yyyy"), f5), 350, 325, 0);



                //Arabic                

                // 1 Serial Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f2), 450, 815, 0);

                // 2 Reference Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.RefNo, f2), 420, 711, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 3 Name
                string ArabicName = absApply.Applicant.Profile.NameAr;
                if (String.IsNullOrEmpty(ArabicName) || ArabicName == "None")
                {
                    ArabicName = absApply.Applicant.Profile.Name;
                }

                float textLength = canvas2.GetEffectiveStringWidth(ArabicName, true);

                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(ArabicName, f2), 420, 696, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4 Date
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(currentDateTime.ToString("dd MMMMMM yyyy", new CultureInfo("ar-QA")), f2), 420, 680, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4.5 Passport Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.Applicant.Profile.IdCardNumber, f2), 420, 665, 0);

                String arabicSemesterName = semester.MonthNameAr + " " + semester.Year.ToString();

                // 5 Semester
                float seTextLength = canvas2.GetEffectiveStringWidth(arabicSemesterName, true);

                float position = ((60 - seTextLength) / 2) + 270;
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_CENTER, new Phrase(arabicSemesterName, f4), 290, 728, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 6 Program
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.Course1.NameAr, f2), 420, 650, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                // 8 Normal Period
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(Convert.ToString(absApply.Course1.CourseDescription.Duration), f2), 420, 633, 0);

                //// 9 Vertual Account Number
                //ColumnText.ShowTextAligned(canvas2,
                //      Element.ALIGN_RIGHT, new Phrase(absApply.HSBCAccPayor, f2), 420, 603, 0);


                //// 9 Classes Commence Date
                //ColumnText.ShowTextAligned(canvas2,
                //      Element.ALIGN_RIGHT, new Phrase(semesterDates.StartDate.Value.ToString("dd MMMMMM yyyy", new CultureInfo("ar-QA")), f2), 460, 326, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                //// 8 Response Date
                //ColumnText.ShowTextAligned(canvas2,
                //      Element.ALIGN_CENTER, new Phrase(currentDateTime.AddDays(14).ToString("dd MMMMMM yyyy", new CultureInfo("ar-QA")), f2), 345, 310, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                ////Third Page Serial Number
                //ColumnText.ShowTextAligned(canvas3,
                //     Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);

                stamper.Close();

            }
            catch (Exception ex)
            {

            }
        }

        public bool SendRejectionToApplicant(AdmissionApply apply, string remarkEn, string remarkAr)
        {
            var subject = "APPLICATION FOR ADMISSION رد على طلب القبول";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <b>Reference :</b>" + apply.RefNo + @"<br />
                                <b>Name :</b>" + apply.ApplicantName + @" <br />
                                <p>Dear Applicant, </p>
                                
                                <p>Assalamualaikom warahmatullah wabarakatuh, </p>

                                <p>We respectfully refer to your application for admission into Al-Madinah International University (MEDIU).</p>
                                <p>The University has thoroughly vetted your application. However, we regret to inform you that your application does not meet our admission requirements. Therefore, your application is <b><font color='#fb0034'>UNSUCCESSFUL</font></b></p>
                                <p>The University is pleased to offer you the admission whenever your application meets the admission requirements.</p>    
                                <p>Note: to find out the reason of the rejection please use this link below</p>
                                <p>http://online.mediu.edu.my/apply/Applicant/replytoofferletter.aspx</p>
                                <p>
                                Thank you
                                </p>
                                <p>                                
                                Al-Madinah International University
                                <br/>
                                Deanship of Admission & Registration
                                </p>

                                    
                            </td>
                            <td width=""50%"" dir=""rtl"">


                            <p><b>الرقم المرجعي:</b>
                            " + apply.RefNo + @"<br /> 
                               <b>الاسم :</b> 
                            " + apply.ApplicantName + @"</p>
                            
                            <p>عزيزي المتقدم/ عزيزتي المتقدمة<br />
                            السلام عليكم ورحمة الله ويركاته
                            </p>
                            
                            <p>نود بكل احترام  الإشارة إلى القرار المتخذ حيال طلبكم المقدم للقبول في جامعة المدينة العالمية.</p>
                            
                            <p>لقد قامت الجامعة بفحص الطلب، ويؤسفنا إبلاغكم بأن شروط القبول بالجامعة لم تتوفر في الطلب،<b><font color='#fb0034'> لذا لم يلق قبولا</font</b>.</p>

                            <p>علما بأن الجامعة مستعدة لقبولكم عندما يكون الطلب مستوفيا للشروط.</p>
                            
                            <p>ملاحظة: يمكنكم الاطلاع على سبب رفض طلبكم في بوابة المتقدمين على هذا الرابط</p>
                            
                            <p>http://online.mediu.edu.my/apply/Applicant/replytoofferletter.aspx?lang=ar</p>
                            <p>
                            شكرا ًجزيلاً
                            </p>
                            
                            <p>
                            جامعة المدينة العالمية <br />
                            عمادة القبول والتسجيل 
                            </p>
                            </td>
                        </tr>
                    </table>
                    ";


            var y = new Random().Next(0, 1000);
            string rejectLetterFileName = apply.Applicant.UserName + "_RejectLetter" + y.ToString() + ".pdf";
            string rejectLetterPath = "";

            string destOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\RejectionLetters\\" + rejectLetterFileName;
            //generate Rejection letter
            rejectLetterPath = GenerateRejectionLetter(apply, destOffer, null);

            if (File.Exists(rejectLetterPath))
            {
                //Send it by Email
                System.IO.BinaryReader br = new System.IO.BinaryReader(System.IO.File.Open(rejectLetterPath, System.IO.FileMode.Open, System.IO.FileAccess.Read));
                br.BaseStream.Position = 0;
                byte[] buffer = br.ReadBytes(Convert.ToInt32(br.BaseStream.Length));
                br.Close();
                //apply.Applicant.Profile.Email
                bool result = svcMessaging.SendEmailWithAttachment(apply.Applicant.Profile.Email, subject, body, true, buffer, rejectLetterFileName);

                //set for review
                apply.LatestOfferLetterFileName = rejectLetterFileName;
                return result;
            }
            else
            {
                return false;
            }
        }

        public string GenerateRejectionLetter(AdmissionApply apply, string offerLetterFileName, DateTime? dateRejectLetterSent)
        {

            string level = apply.Course1.CourseDescription.CourseLevel;
            var y = new Random().Next(0, 1000);
            DateTime currentDateTime = (dateRejectLetterSent != null) ? (DateTime)dateRejectLetterSent : DateTime.Now;
            var s = sm.OpenSession();
            string type = "";

            if ((level != "M") || (level != "P"))
            {
                type = "Postgraduate";
            }
            else
            {
                type = "Undergraduate";
            }


            string sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\Rejection Letter " + type + ".pdf";
            string destOffer = offerLetterFileName;

            try
            {
                IntakeEvent semester = apply.Intake;
                var semesters = s.GetAll<CampusEvent>()
                                           .ThatHasChild(c => c.ParentEvent)
                                               .Where(c => c.Id).IsEqualTo(semester.ParentEvent.Id)
                                           .EndChild()
                                           .And(c => c.EventType).IsEqualTo("SME")
                                           .Execute().ToList();
                CampusEvent semesterDates = semesters.Where(c => c.Code.ToLower().StartsWith(semester.Code.Substring(0, 3).ToLower())).FirstOrDefault();

                AdmissionApply absApply = apply;

                int serial = semester.OfferLetterSerialNumber + 1;
                string cardnum = absApply.Applicant.Profile.IdCardNumber;
                if (cardnum == "nil")
                {
                    cardnum = "";
                }


                PdfReader r = new PdfReader(sourceOffer);
                string guid = Guid.NewGuid().ToString();
                PdfStamper stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate));
                PdfContentByte canvas = stamper.GetOverContent(1);
                PdfContentByte canvas2 = stamper.GetOverContent(2);
                BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                canvas.SetFontAndSize(bf, 8);
                canvas2.SetFontAndSize(bf, 8);

                Font f2 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f3 = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);
                Font f4e = new Font(bf, 8, Font.BOLD, BaseColor.BLACK);
                Font f4 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                Font f4r = new Font(bf, 8, Font.NORMAL, BaseColor.RED);
                Font f5 = new Font(bf, 10, Font.ITALIC, BaseColor.BLACK);

                //English

                // 1 Semester
                //semester.Month
                String semesterName = semester.MonthName.ToUpper() + " " + semester.Year.ToString();

                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_CENTER, new Phrase(semesterName, f5), 288, 727, 0);

                // 2 Serial Number
                string serialNumber = serial.ToString("0000000") + "(" + currentDateTime.Year.ToString().Substring(2, 2) + currentDateTime.Month.ToString("00") + ")";
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f3), 450, 800, 0);

                // 3 Reference Number
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.RefNo, f3), 200, 690, 0);

                // 4 Name
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.NameEnglish.ToUpper(), f3), 200, 674, 0);

                // 5 Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(currentDateTime.ToString("dd.MM.yyyy"), f3), 200, 660, 0);

                //// 6 Passport Number
                //ColumnText.ShowTextAligned(canvas,
                //      Element.ALIGN_LEFT, new Phrase(absApply.Applicant.Profile.IdCardNumber, f3), 200, 659, 0);

                // 7 Program
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(absApply.Course1.NameEn, f3), 200, 646, 0);

                //Arabic                

                // 1 Serial Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_LEFT, new Phrase(serialNumber, f2), 450, 800, 0);

                // 2 Reference Number
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.RefNo, f2), 413, 705, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 3 Name
                string ArabicName = absApply.Applicant.Profile.NameAr;
                if (String.IsNullOrEmpty(ArabicName) || ArabicName == "None")
                {
                    ArabicName = absApply.Applicant.Profile.Name;
                }

                float textLength = canvas2.GetEffectiveStringWidth(ArabicName, true);

                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(ArabicName, f2), 413, 687, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                // 4 Date
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(currentDateTime.ToString("dd MMMMMM yyyy", new CultureInfo("ar-QA")), f2), 413, 668, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                String arabicSemesterName = semester.MonthNameAr + " " + semester.Year.ToString();

                // 5 Semester
                float seTextLength = canvas2.GetEffectiveStringWidth(arabicSemesterName, true);

                float position = ((60 - seTextLength) / 2) + 270;
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_CENTER, new Phrase(arabicSemesterName, f4), 290, 727, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                // 6 Program
                ColumnText.ShowTextAligned(canvas2,
                      Element.ALIGN_RIGHT, new Phrase(absApply.Course1.NameAr, f2), 413, 653, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);


                stamper.Close();

            }
            catch (Exception ex)
            {

            }

            return destOffer;
        }

        public bool SendStudentGraduationApplicationRequest(Student student)
        {
            var body = @"
            <table width = ""100%"" >
            <tr><td>
            <table cellpadding = ""10"" width = ""100%"">
                <tr>
                <td width = ""50%""> 
                    <p>Dear " + student.Profile.Name + @",</p>
                    <p>السلام عليكم و رحمة الله و بركاته</p>
                    <br />
                    <br />
                    
                    <p>Al-Madinah International University (MEDIU) would like to congratulate you for your achievements by completing your program in the University. </p>
                    <br />
                    <br />
                    
                    <p>In preparation for the graduation process, we would like to ask you to fill up the graduation form named: “Graduation Form” is available in student portal in order to unable the University to start certificate preparation process.</p>
                    <br />
                    <br />
                    
                    <p><span style=""color:red""> Important note:</span> Ensure that the information stated in the form is correct and valid. As the university disclaims any error and incorrect information in the form.</p>
                    <br />
                    <br />
                    <p>The attached picture is an explanation of how to get the form in student portal.</p>
                    

                    <p>In the end, we want to note the importance of the initiative to fill out the form listed as certificate issuance procedures depended on it.</p>
                    <br />
                    <br />
                    <p>We wish you a bright future.</p>
                    <br />
                    <br />
                    <p>This letter is for the graduate students only, if you are not graduate student please ignore it.</p>

                    <td style=""text-align: right"" span dir=""rtl"" width = ""50%"">  
                    <p>  عزيزي " + student.Profile.Name + @"  </p>
                    <p>السلام عليكم ورحمة الله وبركاته    </p>
                    <br />
                    <br />
                    <p>تهنئكم جامعة المدينة العالمية (مديو) بماليزيا على إكمالكم لمتطلبات التخرج في البرنامج الذي سجلتم فيه بالجامعة.   </p>
                    <br />
                    <br />
                    <p> وتمهيدا لإجراءات التخرج فإننا نطلب منكم سرعة تعبئة استمارة التخرج  واسمها: استمارة التخرج  الموجودة في بوابة الطالب. حتى تتمكن الجامعة من إصدار وثيقة تخرجك.  </p>
                    <br />
                    <br />
                    <p><span style=""color:red"">  ملاحظة هامة</span>: عليكم التأكد من صحة البيانات المدخلة في الاستمارة خاصة حيث إن الجامعة تخلي مسئوليتها عن أي أخطاء بناءً على هذه الاستمارة.   </p>
                    <br />
                    <br />
                    <p>وقد أرفقنا في المرفق صورة توضيحية لطريقة الحصول على الاستمارة في بوابة الطالب    </p>
                    <br />
                    <br />
                    <p>وفي النهاية نريد أن ننوه بأهمية المبادرة إلى تعبئة الاستمارة المذكورة حيث إن إجراءات إصدار الشهادة تتوقف عليها.   </p>
                    <br />
                    <br />
                    <p> مع تمنياتنا لكم بمستقبل زاهر.  </p>
                    <br />
                    <br />
                    <p> هذه الرسالة للطلبة المتخرجين فقط، فإن لم تكونوا من المتخرجين فنرجو إغفالها.  </p>

                </td>
                </tr>    

           </table>
            </td></tr>
            <tr><td><p><img src=""http://cms.mediu.edu.my/office/content/Images/graduationForm.jpg"" /></p></td></tr>
            </table>
           ";


            return svcMessaging.SendEmail(student.Profile.Email, "Graduation Application Request", body, true);
        }

        public bool SendStudentSvcVerificationLetterEmail(VerificationLetter vl, Student stud)
        {
            var heshe = "";
            var hisher = "";

            var vtype = "";

            if (stud.Profile.Gender == "F")
            {
                heshe = "She";
                hisher = "her";
            }

            else
            {
                heshe = "He";
                hisher = "his";
            }

            if (vl.VLType == "However, the international passport for the above mentioned student already got a student visa valid till ")
            {
                vtype = vl.VLType + vl.VLVisa;
            }
            else
            {
                vtype = vl.VLType;
            }


            var body = @"
                            <table cellpadding = ""10"">
                            <tr>
                                <td class = ""header"" style=""text-align: center"">
                                     <p><img src=""http://cms.mediu.edu.my/office/content/logo_color.jpg"" /></p> <br />
                                </td>
                            </tr>
                            <tr>
                                <td><table>
                                    <tr><td>Date:</td> <td>" + vl.VLDate + @" </td></tr>
                                    <tr><td>Reference No:</td> <td>MEDIU/MY/5.5.5 (" + DateTime.Now.ToString("yy") + ")" + (stud.NumberOfVerificationLetterPerSemester + 1).ToString("D2") + "/" + stud.AdmissionRefNo + @" </td></tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <p><B><U>STUDENT VERIFICATION LETTER</U></B></p> <br>
                                    <p>To: " + vl.VLTo + @" </p>
                                </td>
                            </tr>
                            <tr>
                                <td><table>
                                        <tr>
                                            <td> Student Name : <br /> </td>
                                            <td> <b>" + stud.Profile.Name + @"</b></td>
                                        </tr>
                                        <tr>
                                            <td> Matric No: <br /> </td>
                                            <td> <b>" + stud.MatrixNo + @"</b></td> 
                                        </tr>
                                        <tr>
                                            <td> Nationality : <br /> </td>
                                            <td> <b>" + vl.VLCitizen + @"</b></td> 
                                        </tr>
                                        <tr>
                                            <td> ID No. : <br /> </td>
                                            <td> <b>" + stud.Profile.IdCardNumber + @"</b></td> 
                                        </tr>
                                        <tr>
                                            <td> Programme of Study : <br /> </td>
                                            <td> <b> " + stud.Course.NameEn + @"</b></td> 
                                        </tr>
                                        <tr>
                                            <td> Intake : <br /> </td>
                                            <td> <b>" + stud.IntakeCode + @"</b></td>
                                        </tr>
                                </table>
                                </td>
                            </tr>
                            <tr>
                                <td align=""justify"">
                                    <p>Al-Madinah International University verifies that the above student is an enrolled student who is currently studying 
                                        in the programme " + stud.Course.NameEn + @". " + heshe + @" is a full time student 
                                        from  " + stud.IntakeCode + @" and the normal period of the program is " + stud.Course.CourseDescription.Duration + @"      
                                        " + stud.Course.CourseDescription.DurationType.ToString().ToLower() + @".</p> <br>
                                                                                                                                                                        
                                  
                                    <p>" + vtype + @".</p> <br>
                                    
                                    <p>" + vl.VLAddText + @"</p> <br>
                                             
                                    <p>Kindly, feel free to contact our office if you have further inquiries concerning this student.
                                    This letter has been given upon " + hisher + @" own request.</p>
                                    
                                    <br />
                                    <br />
                                    
                                    <p>Sincerely,</p>                 
                                    <p class=""sign""><img src="""" /></p>
                                    
                                    <p><br />Dean of Students Affairs </p>
                                    <p>“This verification letter valid for 6 months from the date of issue”</p>
                                </td>
                            </tr>

                            <tr>
                                <td align =""left"" valign=""bottom"">
                                <br />
                                <p><img src=""http://cms.mediu.edu.my/office/content/Images/footermediu.jpg"" /></p>
                                </td>
                            </tr>
                        </table>
                        ";
            return svcMessaging.SendEmail(stud.Profile.Email, "Verification Letter", body, true);
        }

        public bool SendNotificationEmail(string emailTo, string subject, string bodyen, string bodyar, Profile personTriggerProfile)
        {
            try
            {
                string recipientEmail;
                recipientEmail = emailTo;

                var body = bodyen + @" <br /> " + bodyar;
                //return svcMessaging.SendEmail(recipientEmail.Replace(";", ","), subject, body, true);
                return svcMessaging.SendEmail(recipientEmail, subject, body, true); //testing for temporary, to check whether the email will be sent or not
            }
            catch (Exception)
            {
                return false;
            }
        }


        public bool SendNotificationEmailUpdateMark(string emailDean, Student student, string username)
        {
            try
            {
                string recipientEmail;
                recipientEmail = emailDean;
                var subject = "Student Mark Changed for " + student.MatrixNo;
                var body = "";
                body = @"<p style='font-size:20px;'><b>Assalammualaikum W.B.T Dean,</b></p>
                                                    <p style='font-size:20px;'><b>Kindly Be Informed That Below Student Subject Mark Has Been Change</p>
                                                    <table style='text-align:left'>
                                                    <tr> 
                                                    <td>Student Name : </td> <td>" + student.Profile.Name + @"</td>
                                                    </tr>
                                                    <tr>
                                                    <td>Student Matrix No : </td> <td>" + student.MatrixNo + @"</td>
                                                    </tr>
                                                    <tr>
                                                   <td>Course Name </td> <td>" + student.CourseName + @"</td>
                                                    </tr>
                                                  <tr>
                                                   <td>Update By </td> <td>" + username + @"</td>
                                                    </tr>
                                                    </table>
                                                    <p>Regards,<br><br>MEDIU</p>";




                //return svcMessaging.SendEmail(recipientEmail.Replace(";", ","), subject, body, true);
                return svcMessaging.SendEmail(recipientEmail, subject, body, true); //testing for temporary, to check whether the email will be sent or not
            }
            catch (Exception)
            {
                return false;
            }
        }


        public bool SendNotificationEmailReSitExam(string emailDean)
        {
            try
            {
                string recipientEmail;
                recipientEmail = emailDean;
                var subject = "Form Re-Sit Exam";
                var body = @"<table width=""100%"">
                    <tr>
                        <td width=""50%"" dir=""ltr"">
                                   <p>Dear Dean,</p>
                                   <p>Kindly been informed that there are Re-Sit exam application need approval in CMS.
                             </td>   

                    </tr>
                       
                                   <p>Thank you</p>
                          
                    </table>
                ";
                return svcMessaging.SendEmail(recipientEmail, subject, body, true); //testing for temporary, to check whether the email will be sent or not
            }
            catch (Exception)
            {
                return false;
            }
        }



        //send email to dean when activate student (enroll applicant)
        public bool SendNotificationEmailEnrolled(string emailDean, AdmissionApply apply, string username)
        {
            try
            {
                string recipientEmail;
                recipientEmail = emailDean;
                var subject = "Activate Student(Enrolled) ";
                var body = @"
                    <table width=""100%"">
                    <tr>
                        <td width=""50%"" dir=""ltr"">
                                   <p>Dear Dean of Registration,</p>
                                   <p>Kindly be informed that below is the applicant that been activate by staff and status changed to Enrolled.
                             </td>   

                    </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <b>Reference Number :</b>" + apply.RefNo + @"<br />
                                <b>Enrolled by :</b>" + username + @"<br />
                             </td>   
                        </tr>
<tr>
                        <td width=""50%"" dir=""ltr"">
                                   <p>Thank you</p>
                                   
                             </td>   

                    </tr>
                    </table>
                    ";


                return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            }
            catch (Exception)
            {
                return false;
            }
        }



        //send email to dean when staff do enforce registration on bulk registration
        public bool SendNotificationEmailBulkRegistration(string emailDean, Student student, string username)
        {
            try
            {
                string recipientEmail;
                recipientEmail = emailDean;
                var subject = "Registration Student Skip Payment Requirement";
                var body = @"
                    <table width=""100%"">
                    <tr>
                        <td width=""50%"" dir=""ltr"">
                                   <p>Dear Dean,</p>
                                   <p>This is the student matrix number that been register by staff through exception system by skip payment requirement in subject registration.
                             </td>   

                    </tr>
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <b>Matrix Number :</b>" + student.MatrixNo + @"<br />
                                <b>Register by :</b>" + username + @"<br />
                             </td>   
                        </tr>
<tr>
                        <td width=""50%"" dir=""ltr"">
                                   <p>Thank you</p>
                                   
                             </td>   

                    </tr>
                    </table>
                    ";


                return svcMessaging.SendEmail(recipientEmail, subject, body, true);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool SendStatementOfGraduationEmail(Student student)
        {
            var subject = "Statement Of Graduation إفادة تخرج";
            var body = @"
                <p dir=""RTL"">
                السلام عليكم ورحمة الله وبركاته
                </p>
                <p dir=""RTL"">
                عزيزي الطالبـ/ــة
                </p>
                <p dir=""RTL"">
                تهانينا: يسر جامعة المدينة العالمية أن تهنئكم لإكمالكم متطلبات البرنامج الدارسي الخاص بكم. 
                ولذا نرفق لكم مع هذه الرسالة إفادة تخرج صالحة لمدة 6 شهور فقط لاستخدامها لحين استخراج الشهادة. 
                </p>
                <p dir=""LTR"">Dear student,</p>
                <p dir=""LTR"">Congratulation, Al-Madinah International University (MEDIU) is pleased that you have fulfilled the requirement for your program. 
                    Therefore, we attached with this letter a statement of graduation which is valid to use for six months up to the issuance of the official certificate.   
                </p>
                <p style=""text-align:center"">
                وتفضلوا بقبول وافر التحية 
                </p>
                <p style=""text-align:center"">
                <br />جامعة المدينة العالمية
                <br />Al-Madinah International University
                <br /><a href=""www.mediu.edu.my"">www.mediu.edu.my</a> 
                </p>               
                    ";
            string statementOfGraduationFileName = student.UserName + ".pdf";
            string statementOfGraduationPath = "";

            string directorypath = @"c:\cms_temp\StatementOfGraduation";
            if (!Directory.Exists(directorypath))
            {
                Directory.CreateDirectory(directorypath);
            }
            string destOffer = directorypath + "\\" + statementOfGraduationFileName;

            // first find the file in filestore. If file not exist, create new file
            var sog_doc = student.Profile.Documents.Where(c => c.Category == "STUDENTRECORD_STATEMENTOFGRADUATION").FirstOrDefault();
            if (sog_doc != null)
            {
                var doc = svcDoc.GetDocument(sog_doc.Id);

                // Send email with statement of graduation pdf file
                bool result = svcMessaging.SendEmailWithAttachment(student.Profile.Email, subject, body, true, doc.Data, doc.Name);
                return result;
            }
            else // create new pdf file for Statement Of Graduation
            {
                if (!File.Exists(destOffer))
                {
                    //generate Rejection letter
                    if (student.IsPostgraduate)
                    {
                        statementOfGraduationPath = GenerateStatementOfGraduationPostGraduate(student, destOffer, null);
                    }
                    else
                    {
                        statementOfGraduationPath = GenerateStatementOfGraduationUnderGraduate(student, destOffer, null);
                    }
                }
                else
                {
                    statementOfGraduationPath = destOffer;
                }
                if (File.Exists(statementOfGraduationPath))
                {
                    //Send it by Email
                    System.IO.BinaryReader br = new System.IO.BinaryReader(System.IO.File.Open(statementOfGraduationPath, System.IO.FileMode.Open, System.IO.FileAccess.Read));
                    br.BaseStream.Position = 0;
                    byte[] buffer = br.ReadBytes(Convert.ToInt32(br.BaseStream.Length));
                    br.Close();

                    // Save the pdf file in filestore
                    var doc = new Mediu.Cms.Domain.Profiles.Document();
                    doc.Data = buffer;
                    doc.Category = "STUDENTRECORD_STATEMENTOFGRADUATION";
                    doc.Name = statementOfGraduationFileName;
                    doc.IsSoftCopyAvailable = true;
                    doc.Title = "Statement Of Graduation";
                    doc.Description = doc.Title + " for " + student.MatrixNo;
                    doc.MimeType = "application/pdf";
                    svcDoc.SaveDocument(student.Profile, doc);

                    // Send email with statement of graduation pdf file
                    bool result = svcMessaging.SendEmailWithAttachment(student.Profile.Email, subject, body, true, buffer, statementOfGraduationFileName);
                    return result;
                }
                else
                {
                    return false;
                }
            }
        }

        private string GenerateStatementOfGraduationUnderGraduate(Student student, string offerLetterFileName, DateTime? dateRejectLetterSent)
        {
            string sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\StatementOfGraduation\\StatementOfGraduationUnderGraduate.pdf";
            string destOffer = offerLetterFileName;
            using (var s = sm.OpenSession())
            {
                try
                {
                    PdfReader r = new PdfReader(sourceOffer);
                    string guid = Guid.NewGuid().ToString();
                    PdfStamper stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate));
                    PdfContentByte canvas = stamper.GetOverContent(1);
                    PdfContentByte canvas2 = stamper.GetOverContent(2);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);

                    Font f2 = new Font(bf, 10, Font.NORMAL, BaseColor.BLACK);

                    // get running no for the statement
                    var statementRunningNo = s.GetAll<GeneralRunningNo>().Where(c => c.Type)
                        .IsEqualTo(GeneralRunningNoType.StatementOfGraduation).Execute().FirstOrDefault();

                    string refNo = "DEAR/ERD/cert/" + DateTime.Now.ToString("yy") + "/" + statementRunningNo.GetCurrentRunningNumber();
                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");

                    // first page in Arabic
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(refNo, f2), 475, 636, 0);
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(currentDate, f2), 475, 619, 0);
                    if (String.IsNullOrEmpty(student.Profile.NameAr))
                    {
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER,
                            new Phrase(student.Profile.Name, f2), 300, 429, 0);
                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER,
                            new Phrase(student.Profile.NameAr, f2), 300, 489, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER,
                        new Phrase(student.Course.NameAr, f2), 300, 440, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                    // second page in English
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_LEFT, new Phrase(refNo, f2), 148, 714, 0);
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_LEFT, new Phrase(currentDate, f2), 148, 698, 0);
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_CENTER, new Phrase(student.Profile.Name, f2), 290, 532, 0);
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_CENTER, new Phrase(student.Course.NameEn, f2), 290, 470, 0);

                    stamper.Close();

                    s.SaveOrUpdate(statementRunningNo);
                    s.Flush();
                }
                catch (Exception ex)
                {
                }
                return destOffer;
            }
        }

        private string GenerateStatementOfGraduationPostGraduate(Student student, string offerLetterFileName, DateTime? dateRejectLetterSent)
        {
            string sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\StatementOfGraduation\\StatementOfGraduationPostGraduate.pdf";
            string destOffer = offerLetterFileName;
            using (var s = sm.OpenSession())
            {
                try
                {
                    PdfReader r = new PdfReader(sourceOffer);
                    string guid = Guid.NewGuid().ToString();
                    PdfStamper stamper = new PdfStamper(r, new FileStream(destOffer, FileMode.OpenOrCreate));
                    PdfContentByte canvas = stamper.GetOverContent(1);
                    PdfContentByte canvas2 = stamper.GetOverContent(2);
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                    canvas.SetFontAndSize(bf, 8);

                    Font f2 = new Font(bf, 10, Font.BOLD, BaseColor.BLACK);
                    Font f3 = new Font(bf, 10, Font.NORMAL, BaseColor.BLUE);
                    Font dateFont = new Font(bf, 12, Font.NORMAL, BaseColor.BLACK);
                    string arabicName = String.IsNullOrEmpty(student.Profile.NameAr) ? student.Profile.Name : student.Profile.NameAr;
                    // get running no for the statement
                    var statementRunningNo = s.GetAll<GeneralRunningNo>().Where(c => c.Type)
                        .IsEqualTo(GeneralRunningNoType.StatementOfGraduation).Execute().FirstOrDefault();

                    string refNo = "DEPS/SGD/cert/" + DateTime.Now.ToString("yy") + "/" + statementRunningNo.GetCurrentRunningNumber();
                    string currentDate = DateTime.Now.ToString("dd/MM/yyyy");
                    string depsEmail = "deps@mediu.edu.my";

                    // first page in Arabic
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(refNo, f2), 474, 620, 0);
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(currentDate, f2), 474, 604, 0);
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_RIGHT, new Phrase(depsEmail, f3), 434, 585, 0);
                    if (String.IsNullOrEmpty(student.Profile.NameAr))
                    {
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER,
                            new Phrase(student.Profile.Name, f2), 300, 444, 0);
                    }
                    else
                    {
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER,
                            new Phrase(student.Profile.NameAr, f2), 300, 440, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);
                    }
                    ColumnText.ShowTextAligned(canvas, Element.ALIGN_CENTER,
                        new Phrase(student.Course.NameAr, f2), 300, 380, 0, PdfWriter.RUN_DIRECTION_RTL, ColumnText.AR_LIG);

                    // second page in English
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_LEFT, new Phrase(refNo, f2), 148, 603, 0);
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_LEFT, new Phrase(currentDate, dateFont), 148, 575, 0);
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_LEFT, new Phrase(depsEmail, f3), 159, 545, 0);
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_CENTER, new Phrase(student.Profile.Name, f2), 300, 390, 0);
                    ColumnText.ShowTextAligned(canvas2, Element.ALIGN_CENTER, new Phrase(student.Course.NameEn, f2), 300, 344, 0);

                    stamper.Close();

                    s.SaveOrUpdate(statementRunningNo);
                    s.Flush();
                }
                catch (Exception ex)
                {
                }
                return destOffer;
            }
        }

        public bool SendDeclarationOfferLetterAcceptance(AdmissionApply apply)
        {
            return SendEmailOfferLetterAcceptance(apply);
        }

        public bool SendEmailOfferLetterAcceptance(AdmissionApply apply)
        {
            var name = apply.Applicant.Profile.Name;
            var refno = apply.RefNo;
            var idno = apply.Applicant.Profile.IdCardNumber;

            var subject = "Declaration Of Offer Letter Acceptance";
            var body = @"
                            <table>
                            <tr>
                                <p style='text-align:center;'><span style='font-weight:bold;'>إقرار موافقة على إشعار القبول</span></p>
                                <p></p>
                                <p>
                                            الإسم            : " + name + @"<br />
                                            الرقم المرجعي   : " + refno + @" <br />
                                            جواز السفر      : " + idno + @" <br /> 
                                </p>
                                <p>أتعهد أنا صاحب البيانات أعلاه بأنني قد اطلعت على إشعار القبول المبدئي للدراسة بجامعة المدينة العالمية وقبلت تنفيذ ما ورد فيه من</p>
                                <p>.شروط والتزامات، وكذلك قمت بالاطلاع على الرسوم الدراسية للبرنامج الدراسي وطرق السداد</p>
                                <p>وهذا إقرار مني بذلك</p>
                                </td></tr>
                                </td></tr>

                                <tr><td dir=""ltr"">
                                <p style='text-align:center;'><span style='font-weight:bold;'>Declaration on offer letter</span></p>


                                Name            : " + name + @"<br />
                                Reference No.   : " + refno + @"<br />
                                Passport No.    : " + idno + @"<br />

                                <p>
                                I declare that I have read all the statements on the offer letter to study in Al-Madinah 
                                International University and I accept to get done all the procedures and rules. Also, I have read  
                                the tuition fees of the program and payment methods. 
                                </p>
                                <p>This is my declaration for so</p>
                                </td></tr></table>";

            if (svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true))
            {
                SaveEmailDeclarationOfferLetter(apply, HttpContext.Current.User.Identity.Name, "Declaration Of Offer Letter Acceptance", body, "Offer Letter Acceptance", "Declaration", "DeclarationOfferLetter");
                return true;
            }
            return false;
        }

        public string CalculateTotalFeesForApplicant(AdmissionApply apply)
        {
            var s = sm.OpenSession();
            string totalFee;
            AdmissionApply absApply = apply;
            var country = s.GetAll<Country>().Where(c => c.Code).IsEqualTo(absApply.Applicant.Profile.Citizenship).Execute().FirstOrDefault();
            string studyMode = absApply.LearningMode.ToString();
            string sourceOffer = "";
            //control Offer Letter Templete
            if (absApply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONCAMPUS")
            {
                if (absApply.Course1.CourseDescription.CourseLevel == "B" || absApply.Course1.CourseDescription.CourseLevel == "D")
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Undergraduate_auto.pdf";
                else
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Postgraduate_auto.pdf";
            }
            else
            {
                if (absApply.Course1.CourseDescription.CourseLevel == "B" || absApply.Course1.CourseDescription.CourseLevel == "D")
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Undergraduate_auto_MY.pdf";
                else
                    sourceOffer = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\OfferLetters\\Offer_Letter_Postgraduate_auto_MY.pdf";
            }

            try
            {

                IntakeEvent semester = apply.Intake;
                var semesters = s.GetAll<CampusEvent>()
                                            .ThatHasChild(c => c.ParentEvent)
                                                .Where(c => c.Id).IsEqualTo(semester.ParentEvent.Id)
                                            .EndChild()
                                            .And(c => c.EventType).IsEqualTo("SME")
                                            .Execute().ToList();
                CampusEvent semesterDates = semesters.Where(c => c.Code.ToLower().StartsWith(semester.Code.Substring(0, 3).ToLower())).FirstOrDefault();

                string yearAr = "";
                string monthAr = "";
                string semesterIntake = semester.MonthName.ToUpper() + " " + semester.Year.ToString();


                //Calculate admin fee
                var immigrationBondFee = s.GetAll<ImmigrationBondFeeGroup>().Where(c => c.ImmigrationCountryGroup).IsEqualTo(country.ImmigrationCountryGroup).Execute().FirstOrDefault();
                List<FeeStructureItem> applicantFeeStructureItems = new List<FeeStructureItem>();
                var courseGroupItem = s.GetOne<CourseGroupItem>().Where(x => x.Course).IsEqualTo(apply.Course1)
                    .AndHasChild(c => c.CourseGroup).Where(x => x.IsActive).IsEqualTo(true).EndChild().Execute();
                if (courseGroupItem != null)
                {
                    var feeStructure = s.GetOne<FeeStructure>().Where(x => x.CourseGroup).IsEqualTo(courseGroupItem.CourseGroup).Execute();
                    foreach (var item in feeStructure.FeeStructureItems.Where(x => x.Stage == Stage.One))
                    {
                        if (apply.LearningMode == LearningMode.OnCampus)
                        {
                            if (apply.Applicant.Profile.Citizenship == "MY" || studyMode.ToUpper() == "ONLINE")
                            {
                                if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.Local || item.StudentType == StudentType.Both))
                                {
                                    applicantFeeStructureItems.Add(item);
                                }
                            }
                            else
                            {
                                if ((item.Mode == Mode.OnCampus || item.Mode == Mode.Both) && (item.StudentType == StudentType.International || item.StudentType == StudentType.Both))
                                {

                                    //Added to dinamicly get the Amount for ImmigrationBondGroup
                                    if (item.Mode == Mode.OnCampus && item.StudentType == StudentType.International && item.RevenueCategory == RevenueCategory.ImmigrationBond && item.Stage == Stage.One)
                                    {
                                        item.AmountMYR = immigrationBondFee.AmountMYR;
                                        item.AmountUSD = immigrationBondFee.AmountUSD;
                                        applicantFeeStructureItems.Add(item);
                                    }
                                    else
                                    {
                                        applicantFeeStructureItems.Add(item);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (item.Mode == Mode.Online || item.Mode == Mode.Both)
                            {
                                applicantFeeStructureItems.Add(item);
                            }
                        }
                    }
                }


                //--- Set additional fees

                string immiBond = immigrationBondFee.AmountMYR.ToString("0");
                List<FeeStructureItem> othersFees = new List<FeeStructureItem>()
                                                         {new FeeStructureItem() { Id=new Guid(), Name = "MAPT/MEPT Fee", NameAr = "امتحان الكفاءة في اللغة العربية/ الإنجليزية", AmountMYR = 100 },
                                                          //new FeeStructureItem() { Id=new Guid(), Name = "Calling Visa Fee", NameAr = "رسوم الفيزا" , AmountMYR = 370}, 
                                                          //new FeeStructureItem() { Id=new Guid(), Name = "Insurance Fee", NameAr = "رسوم التأمين الصحي" , AmountMYR = 350}, 
                                                          new FeeStructureItem() { Id=new Guid(), Name = "University Bond Fee", NameAr = "ضمان الجامعة" , AmountMYR = 1000} 
                                                         //new FeeStructureItem() { Id=new Guid(), Name = "Immigration Bond Fee", NameAr = "رسوم تأمين الجوازات" , AmountMYR = immigrationBondFee.AmountMYR}
                                                          
                                                          };
                //English


                // 1 Semester
                string semesterName = "ADMISSION FOR " + semesterIntake + " ( " + studyMode.ToUpper() + " )";


                var dur = absApply.Course1.CourseDescription.DurationType;
                string durationType = "";
                string durationTypeAr = "";
                if (dur == DurationType.Months)
                {
                    durationType = "Months";
                    durationTypeAr = "(أشهر)";
                }
                else if (dur == DurationType.Weeks)
                {
                    durationType = "Weeks";
                    durationTypeAr = "(أسابيع)";
                }
                else
                {
                    durationType = "Years";
                    durationTypeAr = "(سنة/سنوات)";
                }

                //Extra Requirement List 
                int erListIndexEn = 604;
                int erListIndexAr = 585;



                //All fees
                int i = 0;
                decimal totalMyr = 0;

                if (applicantFeeStructureItems != null)
                    foreach (var fs in applicantFeeStructureItems.OrderBy(c => c.Name))
                    {

                        ++i;
                        totalMyr = totalMyr + Convert.ToDecimal(fs.AmountMYR);
                    }

                // Add immigration bond, calling visa and insurance
                if (apply.Applicant.Profile.Citizenship != "MY" && studyMode.ToUpper() == "ONCAMPUS")
                {

                    if (!absApply.PlacementTestExamRequired)
                    {
                        var itemToRemove = othersFees.Where(c => c.Name.Contains("MAPT/MEPT Fee")).LastOrDefault();
                        othersFees.Remove(itemToRemove);
                    }

                    foreach (var fee in othersFees)
                    {
                        ++i;
                        totalMyr = totalMyr + Convert.ToDecimal(fee.AmountMYR);
                    }

                }

                //TOTAL FEES         




                if (studyMode.ToUpper() == "ONLINE" && absApply.Applicant.Profile.Citizenship != "MY")
                // if (studyMode.ToUpper() == "ONLINE")
                {

                    // totalMyr = 2000;
                    totalMyr = 3000;

                }

                if ((studyMode.ToUpper() == "ONCAMPUS" && absApply.Applicant.Profile.Citizenship != "MY") && totalMyr < 5000)
                {

                    totalMyr = 5000;

                }

                totalFee = totalMyr.ToString("0.00");
            }
            catch (Exception ex)
            {
                return "";
            }

            return totalFee;
        }

        //Ahmad Al-Ahmad
        public bool SendEmailAfterOfferLetterAcceptancePaymentReminder(AdmissionApply apply)
        {
            var name = apply.Applicant.Profile.Name;
            var refno = apply.RefNo;
            // string totalMyr = CalculateTotalFeesForApplicant(apply);
            string totalMyr = "0.00";
            var idno = apply.Applicant.Profile.IdCardNumber;
            //0==> online 
            var appLearningMode = apply.Applicant.AdmissionApply.LearningMode.ToString();

            var subject = "طلب سداد الرسوم الإدارية Administrative fees request payment ";

            var body = "";
            if (appLearningMode.ToUpper() == "ONLINE" || (appLearningMode.ToUpper() == "ONCAMPUS" && apply.Applicant.Profile.Citizenship == "MY"))
            {
                body = @"
<body>
  <style type='text/css'>

 table.MsoNormalTable
	{line-height:115%;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
 p.MsoNormal
	{margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:'Times New Roman','serif';
	        margin-left: 0in;
            margin-right: 0in;
            margin-top: 0in;
        }
p
	{margin-right:0in;
	margin-left:0in;
	font-size:12.0pt;
	font-family:'Times New Roman','serif';
	}
a:link
	{color:blue;
	text-decoration:underline;
	text-underline:single;
        }
p.MsoListParagraphCxSpFirst
	{margin-top:0in;
	margin-right:.5in;
	margin-bottom:0in;
	margin-left:0in;
	margin-bottom:.0001pt;
	text-align:right;
	line-height:115%;
	direction:rtl;
	unicode-bidi:embed;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
p.MsoListParagraphCxSpMiddle
	{margin-top:0in;
	margin-right:.5in;
	margin-bottom:0in;
	margin-left:0in;
	margin-bottom:.0001pt;
	text-align:right;
	line-height:115%;
	direction:rtl;
	unicode-bidi:embed;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;
        }
p.MsoListParagraphCxSpLast
	{margin-top:0in;
	margin-right:.5in;
	margin-bottom:10.0pt;
	margin-left:0in;
	text-align:right;
	line-height:115%;
	direction:rtl;
	unicode-bidi:embed;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
    </style>
    <div align='right'>
        <table border='0' cellpadding='0' cellspacing='0' class='MsoNormalTable' 
            dir='rtl' style='border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0in 0in 0in 0in;
 mso-table-dir:bidi'>
            <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
                <td style='width:239.4pt;border:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt' valign='top' width='319'>
                    <p align='center' class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;
  text-align:center;direction:rtl;unicode-bidi:embed'>
                        <b><span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;color:red'>
                        طلب سداد الرسوم الإدارية</span></b><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='margin-top:6.0pt;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>عزيزي 
                        الطالبـــ/ ـــة</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='margin-top:6.0pt;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>الرقم المرجعي: ";
                body += refno;
                body += @"</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='margin-top:6.0pt;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>الاسم:";
                body += name;
                body += @"</span><span lang='AR-EG' style='mso-bidi-language:AR-EG'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <b>
                        <span lang='AR-SA' style='font-family:
  &quot;Traditional Arabic&quot;,&quot;serif&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        السلام عليكم ورحمة الله وبركاته<o:p></o:p></span></b></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;تهانينا!! تود جامعة المدينة العالمية بماليزيا إشعاركم بأنه قد تم 
                        قبولكم مبدئيا للدراسة بجامعة المدينة العالمية</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:justify;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        ولاستكمال باقي الإجراءات المتعلقة بقبولكم النهائي وتمكينكم من الدخول إلى بوابة 
                        الطالب وتحويل وضعكم إلى طالب نشط عليكم سرعة المبادرة إلى سداد قيمة الرسوم 
                        المذكورة في إشعار القبول والمبدئي</span><span dir='LTR' lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'> </span>
                        <span lang='AR-EG' style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>وهي <span style='background: yellow; mso-highlight: yellow'>";
                body += totalMyr;
                body += @" </span> رنجت 
                        ماليزى</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:justify;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>فنأمل 
                        منكم سرعة المبادرة بالتسديد، لنتمكن من تقديم خدماتنا إليكم، وفي حال عدم السداد 
                        فإن الجامعة تأسف لعدم إكمال باقي إجراءات قبولكم.</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p dir='RTL' style='margin-bottom:6.0pt;text-align:justify;direction:rtl;
  unicode-bidi:embed'>
                        <b>
                        <span lang='AR-EG' style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  color:red;mso-bidi-language:AR-EG'>ولمعرفة طرق السداد يرجى اتباع الرابط التالي: </span></b>
                        <span lang='AR-SA'><o:p></o:p></span>
                    </p>
                    <p align='right' class='MsoNormal' dir='LTR' style='mso-margin-top-alt:auto;
  mso-margin-bottom-alt:auto;text-align:right'>
                        <a href='http://www.mediu.edu.my/ar/?page_id=30900' target='_blank'>
                        http://www.mediu.edu.my/ar/?page_id=30900</a><span dir='RTL' 
                            lang='AR-SA'><o:p></o:p></span></p>
                    <p dir='RTL' style='margin-bottom:6.0pt;text-align:justify;direction:rtl;
  unicode-bidi:embed'>
                        <b>
                        <span lang='AR-EG' style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>لمعرفة تفاصيل المبلغ المطلوب يرجى اتباع إحدى الطرق الآتية: </span></b>
                        <span lang='AR-SA'><o:p></o:p></span>
                    </p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>1</span><span 
                            lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>.</span><span 
                            lang='AR-SA' style='font-size:11.0pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>مراجعة إشعار القبول المبدئي الذي أرسل لكم سابقًا. </span>
                    </p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        2.</span><span lang='AR-SA' style='font-size:11.0pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>أو الدخول إلى </span>
                        <a href='http://online.mediu.edu.my/apply/applicant/login.aspx?lango=ar&amp;lango=en&amp;lango=ar&amp;lango=en&amp;lango=ar&amp;lang=en' 
                            target='_blank'><b>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>بوابة المتقدمين</span></b></a><span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;mso-bidi-language:
  AR-EG'> باستخدام بيانات الدخول الخاصة بكم، ومتابعة تسلسل إجراءات الطلب الخاص بكم والملاحظات المسجلة.</span><span 
                            lang='AR-SA'><o:p></o:p></span></p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        3.</span><span lang='AR-SA' style='font-size:11.0pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>أو الاتصال المباشر على المركز التعليمي الذي تتبعونه. </span>
                        <span lang='AR-SA'><o:p></o:p></span>
                    </p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>أو الدخول إلى المحادثة المباشرة على موقع الجامعة.</span><a 
                            href='http://www.mediu.edu.my' target='_blank'><span dir='LTR' 
                            style='font-size:11.0pt'>www.mediu.edu.my</span></a><span dir='LTR' 
                            style='font-size:11.0pt'> &nbsp;</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;mso-margin-bottom-alt:
  auto;text-align:justify;text-indent:.5in;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        شاكرين لكم سلفًا ومقدرين حسن استجابتكم. </span><span lang='AR-SA'><o:p></o:p>
                        </span>
                    </p>
                </td>
                <td style='width:239.4pt;border:solid windowtext 1.0pt;
  border-right:none;padding:0in 5.4pt 0in 5.4pt' valign='top' width='319'>
                    <p align='center' class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:
  center'>
                        <span style='font-size:13.0pt;color:red'>Administrative fees request payment<span 
                            dir='RTL' lang='AR-SA'><o:p></o:p></span></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>Dear Student<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>Reference No:";
                body += refno;
                body += @"<span dir='RTL' lang='AR-SA'>...<o:p></o:p></span></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>Name:";
                body += name;
                body += @"<span dir='RTL' lang='AR-SA'></span><o:p></o:p></span></p>
                    <p align='center' class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:
  center'>
                        <b>
                        <span style='font-size:13.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'><o:p>
                        &nbsp;</o:p></span></b></p>
                    <p align='center' class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:
  center'>
                        <b><span dir='RTL' lang='AR-SA' 
                            style='font-size:13.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        السلام عليكم ورحمة الله وبركاته</span><span style='font-size:13.0pt;
  font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'><o:p></o:p></span></b></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt'><o:p>&nbsp;</o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>&nbsp;</span><span style='font-size:13.0pt;
  font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span>Congratulations!! Al-Madinah International University (MEDIU) is pleased 
                        to inform you that you have been accepted as a student in Al-Madinah 
                        International University.<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'>However, in order for us to complete the rest of the procedures for the registration in order 
                        to enable you to access student’s portal. To be enrolled as an active student, 
                        you are required to pay <span style='background: yellow; mso-highlight: yellow'>
                        RM";
                body += totalMyr;
                body += @"</span>.<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>In case you didn&#39;t settle 
                        your payment in the exact time, unfortunately the university will not precede 
                        with your acceptance procedures.<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <b>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;color:red;
  mso-bidi-language:AR-EG'>To find out about our payment methods please follow this link:<o:p></o:p></span></b></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <a href='http://www.mediu.edu.my/admissions/payment-methods.html' target='_blank'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'>
                        http://www.mediu.edu.my/admissions/payment-methods.html</span></a><span 
                            style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>For more information about 
                        your payment details, you can check one of the following:<o:p></o:p></span></p>
                    <p class='MsoListParagraphCxSpFirst' dir='LTR' style='margin-top:6.0pt;
  margin-right:0in;margin-bottom:0in;margin-left:0in;margin-bottom:.0001pt;
  mso-add-space:auto;text-align:left;text-indent:-.25in;line-height:normal;
  mso-list:l0 level1 lfo1;direction:ltr;unicode-bidi:embed'>
                        <![if !supportLists]>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-fareast-font-family:
  &quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-list:Ignore'>1.<span 
                            style='font:7.0pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span></span></span><![endif]>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>Review the offer letter, 
                        which has sent you earlier.<o:p></o:p></span></p>
                    <p class='MsoListParagraphCxSpMiddle' dir='LTR' style='margin-top:6.0pt;
  margin-right:0in;margin-bottom:0in;margin-left:0in;margin-bottom:.0001pt;
  mso-add-space:auto;text-align:left;text-indent:-.25in;line-height:normal;
  mso-list:l0 level1 lfo1;direction:ltr;unicode-bidi:embed'>
                        <![if !supportLists]>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-fareast-font-family:
  &quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-list:Ignore'>2.<span 
                            style='font:7.0pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span></span></span><![endif]>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>Log in to </span>
                        <a href='http://online.mediu.edu.my/apply/applicant/login.aspx?lango=ar&amp;lango=en&amp;lango=ar&amp;lango=en&amp;lango=ar&amp;lang=en' 
                            target='_blank'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'>
                        applicants portal</span></a><span class='MsoHyperlink'><span style='font-size:13.0pt;
  font-family:&quot;Arabic Typesetting&quot;;text-decoration:none;text-underline:none'> </span></span>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'>using your login details and follow the sequence of your application procedures and the 
                        recorded notes.<o:p></o:p></span></p>
                    <p class='MsoListParagraphCxSpLast' dir='LTR' style='margin-top:6.0pt;margin-right:
  0in;margin-bottom:0in;margin-left:0in;margin-bottom:.0001pt;mso-add-space:
  auto;text-align:left;text-indent:-.25in;line-height:normal;mso-list:l0 level1 lfo1;
  direction:ltr;unicode-bidi:embed'>
                        <![if !supportLists]>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-fareast-font-family:
  &quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-list:Ignore'>3.<span 
                            style='font:7.0pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span></span></span><![endif]>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>Contact with your 
                        correspondence Virtual center, or with our customer service center Via Live 
                        Chat through website </span><a href='http://www.mediu.edu.my' target='_blank'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'>
                        www.mediu.edu.my</span></a><span 
                            style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'> </span>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'><o:p></o:p></span>
                    </p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;'>Your co-operation is highly appreciated<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;'>Thank you</span><span dir='RTL' lang='AR-EG' style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'><o:p></o:p></span></p>
                </td>
            </tr>
        </table>
    </div>

</body>";
            }
            else
            {
                body = @"
<body>
  <style type='text/css'>

 table.MsoNormalTable
	{line-height:115%;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
 p.MsoNormal
	{margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:'Times New Roman','serif';
	        margin-left: 0in;
            margin-right: 0in;
            margin-top: 0in;
        }
p
	{margin-right:0in;
	margin-left:0in;
	font-size:12.0pt;
	font-family:'Times New Roman','serif';
	}
a:link
	{color:blue;
	text-decoration:underline;
	text-underline:single;
        }
p.MsoListParagraphCxSpFirst
	{margin-top:0in;
	margin-right:.5in;
	margin-bottom:0in;
	margin-left:0in;
	margin-bottom:.0001pt;
	text-align:right;
	line-height:115%;
	direction:rtl;
	unicode-bidi:embed;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
p.MsoListParagraphCxSpMiddle
	{margin-top:0in;
	margin-right:.5in;
	margin-bottom:0in;
	margin-left:0in;
	margin-bottom:.0001pt;
	text-align:right;
	line-height:115%;
	direction:rtl;
	unicode-bidi:embed;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;
        }
p.MsoListParagraphCxSpLast
	{margin-top:0in;
	margin-right:.5in;
	margin-bottom:10.0pt;
	margin-left:0in;
	text-align:right;
	line-height:115%;
	direction:rtl;
	unicode-bidi:embed;
	font-size:11.0pt;
	font-family:'Calibri','sans-serif';
	}
    </style>
    <div align='right'>
        <table border='0' cellpadding='0' cellspacing='0' class='MsoNormalTable' 
            dir='rtl' style='border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0in 0in 0in 0in;
 mso-table-dir:bidi'>
            <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
                <td style='width:239.4pt;border:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt' valign='top' width='319'>
                    <p align='center' class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;
  text-align:center;direction:rtl;unicode-bidi:embed'>
                        <b><span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;color:red'>
                        طلب سداد الرسوم الإدارية</span></b><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='margin-top:6.0pt;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>عزيزي 
                        الطالبـــ/ ـــة</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='margin-top:6.0pt;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>الرقم المرجعي: ";
                body += refno;
                body += @"</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='margin-top:6.0pt;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>الاسم:";
                body += name;
                body += @"</span><span lang='AR-EG' style='mso-bidi-language:AR-EG'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <b>
                        <span lang='AR-SA' style='font-family:
  &quot;Traditional Arabic&quot;,&quot;serif&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        السلام عليكم ورحمة الله وبركاته<o:p></o:p></span></b></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:right;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;تهانينا!! تود جامعة المدينة العالمية بماليزيا إشعاركم بأنه قد تم 
                        قبولكم مبدئيا للدراسة بجامعة المدينة العالمية</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:justify;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        ولاستكمال باقي الإجراءات المتعلقة باستخراج تأشيرة الاستقدام الى ماليزيا و قبولكم النهائي وتمكينكم من الدخول إلى بوابة 
                        الطالب وتحويل وضعكم إلى طالب نشط عليكم سرعة المبادرة إلى سداد قيمة الرسوم 
                        المذكورة في إشعار القبول والمبدئي</span><span dir='LTR' lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'> </span>
                        <span lang='AR-EG' style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>وهي <span style='background: yellow; mso-highlight: yellow'>3000</span> رنجت 
                        ماليزى</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;text-align:justify;
  direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>فنأمل 
                        منكم سرعة المبادرة بالتسديد، لنتمكن من تقديم خدماتنا إليكم، وفي حال عدم السداد 
                        فإن الجامعة تأسف لعدم إكمال باقي إجراءات قبولكم.</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p dir='RTL' style='margin-bottom:6.0pt;text-align:justify;direction:rtl;
  unicode-bidi:embed'>
                        <b>
                        <span lang='AR-EG' style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  color:red;mso-bidi-language:AR-EG'>ولمعرفة طرق السداد يرجى اتباع الرابط التالي: </span></b>
                        <span lang='AR-SA'><o:p></o:p></span>
                    </p>
                    <p align='right' class='MsoNormal' dir='LTR' style='mso-margin-top-alt:auto;
  mso-margin-bottom-alt:auto;text-align:right'>
                        <a href='http://www.mediu.edu.my/ar/?page_id=30900' target='_blank'>
                        http://www.mediu.edu.my/ar/?page_id=30900</a><span dir='RTL' 
                            lang='AR-SA'><o:p></o:p></span></p>
                    <p dir='RTL' style='margin-bottom:6.0pt;text-align:justify;direction:rtl;
  unicode-bidi:embed'>
                        <b>
                        <span lang='AR-EG' style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>لمعرفة تفاصيل المبلغ المطلوب يرجى اتباع إحدى الطرق الآتية: </span></b>
                        <span lang='AR-SA'><o:p></o:p></span>
                    </p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>1</span><span 
                            lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>.</span><span 
                            lang='AR-SA' style='font-size:11.0pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>مراجعة إشعار القبول المبدئي الذي أرسل لكم سابقًا. </span>
                    </p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        2.</span><span lang='AR-SA' style='font-size:11.0pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>أو الدخول إلى </span>
                        <a href='http://online.mediu.edu.my/apply/applicant/login.aspx?lango=ar&amp;lango=en&amp;lango=ar&amp;lango=en&amp;lango=ar&amp;lang=en' 
                            target='_blank'><b>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>بوابة المتقدمين</span></b></a><span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;mso-bidi-language:
  AR-EG'> باستخدام بيانات الدخول الخاصة بكم، ومتابعة تسلسل إجراءات الطلب الخاص بكم والملاحظات المسجلة.</span><span 
                            lang='AR-SA'><o:p></o:p></span></p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        3.</span><span lang='AR-SA' style='font-size:11.0pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>أو الاتصال المباشر على المركز التعليمي الذي تتبعونه. </span>
                        <span lang='AR-SA'><o:p></o:p></span>
                    </p>
                    <p dir='RTL' style='margin-top:5.0pt;margin-right:.5in;margin-bottom:6.0pt;
  margin-left:0in;text-align:justify;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-EG' style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;;
  mso-bidi-language:AR-EG'>أو الدخول إلى المحادثة المباشرة على موقع الجامعة.</span><a 
                            href='http://www.mediu.edu.my' target='_blank'><span dir='LTR' 
                            style='font-size:11.0pt'>www.mediu.edu.my</span></a><span dir='LTR' 
                            style='font-size:11.0pt'> &nbsp;</span><span lang='AR-SA'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='RTL' style='mso-margin-top-alt:auto;mso-margin-bottom-alt:
  auto;text-align:justify;text-indent:.5in;direction:rtl;unicode-bidi:embed'>
                        <span lang='AR-SA' 
                            style='font-size:11.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        شاكرين لكم سلفًا ومقدرين حسن استجابتكم. </span><span lang='AR-SA'><o:p></o:p>
                        </span>
                    </p>
                </td>
                <td style='width:239.4pt;border:solid windowtext 1.0pt;
  border-right:none;padding:0in 5.4pt 0in 5.4pt' valign='top' width='319'>
                    <p align='center' class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:
  center'>
                        <span style='font-size:13.0pt;color:red'>Administrative fees request payment<span 
                            dir='RTL' lang='AR-SA'><o:p></o:p></span></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>Dear Student<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>Reference No:";
                body += refno;
                body += @"<span dir='RTL' lang='AR-SA'>...<o:p></o:p></span></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>Name:";
                body += name;
                body += @"<span dir='RTL' lang='AR-SA'></span><o:p></o:p></span></p>
                    <p align='center' class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:
  center'>
                        <b>
                        <span style='font-size:13.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'><o:p>
                        &nbsp;</o:p></span></b></p>
                    <p align='center' class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:
  center'>
                        <b><span dir='RTL' lang='AR-SA' 
                            style='font-size:13.0pt;font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'>
                        السلام عليكم ورحمة الله وبركاته</span><span style='font-size:13.0pt;
  font-family:&quot;Traditional Arabic&quot;,&quot;serif&quot;'><o:p></o:p></span></b></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt'><o:p>&nbsp;</o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt'>&nbsp;</span><span style='font-size:13.0pt;
  font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span>Congratulations!! Al-Madinah International University (MEDIU) is pleased 
                        to inform you that you have been accepted as a student in Al-Madinah 
                        International University.<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt;text-align:justify'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'>However, in order for us to complete the rest of the procedures for the extraction of calleing visa to Malaysia and final addmission to enable you to access student’s portal. To be enrolled as an active student, 
                        you are required to pay <span style='background: yellow; mso-highlight: yellow'>
                        RM3000</span>.<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>In case you didn&#39;t settle 
                        your payment in the exact time, unfortunately the university will not precede 
                        with your acceptance procedures.<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <b>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;color:red;
  mso-bidi-language:AR-EG'>To find out about our payment methods please follow this link:<o:p></o:p></span></b></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <a href='http://www.mediu.edu.my/admissions/payment-methods.html' target='_blank'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'>
                        http://www.mediu.edu.my/admissions/payment-methods.html</span></a><span 
                            style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'><o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>For more information about 
                        your payment details, you can check one of the following:<o:p></o:p></span></p>
                    <p class='MsoListParagraphCxSpFirst' dir='LTR' style='margin-top:6.0pt;
  margin-right:0in;margin-bottom:0in;margin-left:0in;margin-bottom:.0001pt;
  mso-add-space:auto;text-align:left;text-indent:-.25in;line-height:normal;
  mso-list:l0 level1 lfo1;direction:ltr;unicode-bidi:embed'>
                        <![if !supportLists]>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-fareast-font-family:
  &quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-list:Ignore'>1.<span 
                            style='font:7.0pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span></span></span><![endif]>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>Review the offer letter, 
                        which has sent you earlier.<o:p></o:p></span></p>
                    <p class='MsoListParagraphCxSpMiddle' dir='LTR' style='margin-top:6.0pt;
  margin-right:0in;margin-bottom:0in;margin-left:0in;margin-bottom:.0001pt;
  mso-add-space:auto;text-align:left;text-indent:-.25in;line-height:normal;
  mso-list:l0 level1 lfo1;direction:ltr;unicode-bidi:embed'>
                        <![if !supportLists]>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-fareast-font-family:
  &quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-list:Ignore'>2.<span 
                            style='font:7.0pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span></span></span><![endif]>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>Log in to </span>
                        <a href='http://online.mediu.edu.my/apply/applicant/login.aspx?lango=ar&amp;lango=en&amp;lango=ar&amp;lango=en&amp;lango=ar&amp;lang=en' 
                            target='_blank'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'>
                        applicants portal</span></a><span class='MsoHyperlink'><span style='font-size:13.0pt;
  font-family:&quot;Arabic Typesetting&quot;;text-decoration:none;text-underline:none'> </span></span>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'>using your login details and follow the sequence of your application procedures and the 
                        recorded notes.<o:p></o:p></span></p>
                    <p class='MsoListParagraphCxSpLast' dir='LTR' style='margin-top:6.0pt;margin-right:
  0in;margin-bottom:0in;margin-left:0in;margin-bottom:.0001pt;mso-add-space:
  auto;text-align:left;text-indent:-.25in;line-height:normal;mso-list:l0 level1 lfo1;
  direction:ltr;unicode-bidi:embed'>
                        <![if !supportLists]>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-fareast-font-family:
  &quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'><span style='mso-list:Ignore'>3.<span 
                            style='font:7.0pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        </span></span></span><![endif]>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:AR-EG'>Contact with your 
                        correspondence Virtual center, or with our customer service center Via Live 
                        Chat through website </span><a href='http://www.mediu.edu.my' target='_blank'>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'>
                        www.mediu.edu.my</span></a><span 
                            style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;'> </span>
                        <span style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'><o:p></o:p></span>
                    </p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;'>Your co-operation is highly appreciated<o:p></o:p></span></p>
                    <p class='MsoNormal' dir='LTR' style='margin-top:6.0pt'>
                        <span style='font-size:
  13.0pt;font-family:&quot;Arabic Typesetting&quot;'>Thank you</span><span dir='RTL' lang='AR-EG' style='font-size:13.0pt;font-family:&quot;Arabic Typesetting&quot;;mso-bidi-language:
  AR-EG'><o:p></o:p></span></p>
                </td>
            </tr>
        </table>
    </div>

</body>
";
            }


            if (svcMessaging.SendEmail(apply.Applicant.Profile.Email, subject, body, true))
            {
                return true;
            }
            //if (svcMessaging.SendEmail("ahmad.ahmad@mediu.edu.my", subject, body, true))
            //{
            //    return true;
            //}
            return false;
        }

        private void SaveEmailDeclarationOfferLetter(AdmissionApply receiver, string sender, string description, string body, string category, string title, string filename)
        {
            var emailDocument = new Mediu.Cms.Domain.Profiles.EmailDocument();
            emailDocument.Data = CreateDocumentForEmail(receiver.Applicant.Profile.Email, sender, title, body);
            emailDocument.MimeType = "text/html";
            emailDocument.IsSoftCopyAvailable = true;
            emailDocument.Title = title;
            emailDocument.Description = description;
            emailDocument.Name = filename;
            emailDocument.Category = category;
            receiver.Applicant.Profile.AddEmailDocument(emailDocument);
            emailDocumentSvc.SaveDocument(receiver.Applicant.Profile, emailDocument);
        }

        public bool SendFirstWarningLowCgpa(Student student, string semesterCode)
        {
            if (student.Course.CourseDescription.CourseLevel == "B")
            {
                return SendFirstWarningLowCgpaBachelor(student, semesterCode);
            }
            else if (student.Course.IsPostGraduate && !student.Course.IsResearch)
            {
                return SendFirstWarningLowCgpaPostgraduate(student, semesterCode);
            }
            return false;
        }

        public bool SendFirstWarningLowCgpaBachelor(Student student, string semesterCode)
        {
            string nameAr = String.IsNullOrEmpty(student.Profile.NameAr) ? student.Profile.Name : student.Profile.NameAr;
            var subject = "The First Warning Letter due to Low Cumulative Grade Point Average (CGPA)";
            var body = @"
                            <table>
                            <tr>
                                <td class = ""header"" style=""text-align: center"">
                                             <p><a href=""www.mediu.edu.my""><img src=""http://cms.mediu.edu.my/office/content/logo_color.jpg"" /></a></p> <br />
                                        </td>
                                    </tr>
                                    <tr style='padding-bottom:10px;'><td dir=""rtl"">
                                    <p style='text-align:center;'><span style='font-weight:bold;'>خطاب الإنذار الأول بسبب انخفاض المعدل التراكمي - البكالوريوس</span></p>
                                    <p>عزيزي الطالب/ عزيزي الطالبة؛</p>
                                    <p>
                                                الإسم : " + nameAr + @"<br />
                                                الرقم المرجعي : " + student.MatrixNo + @" <br />
                                                التاريخ : <span dir=""ltr"">" + DateTime.Now.ToString("dd MMMM yyyy") + @"</span><br />
                                    </p>
                                    <p>السلام عليكم ورحمة الله وبركاته</p>
                                    <p>نظرا لانخفاض معدلكم التراكمي عن أقل نسبة للنجاح (2.00) للمرة الأولى؛ فإن قسم الامتحانات والسجلات بعمادة القبول والتسجيل يوجه لكم هذا الخطاب تنبيها وحثا لكم على ضرورة بذل الجهد والطاقة في رفع معدلكم التراكمي ليصبح (2.00) فما فوق في الفصل التالي، ويتأسف القسم في إلغاء قيدكم (فصلكم) من الدراسة في حالة عجزكم عن تحقيق ذلك، علما بأنه في حالة الفصل لا يمكنكم الدراسة بالجامعة مرة أخرى إلا بتقديم طلب جديد.</p>
                                    <p>مع التمنيات لكم بالتوفيق والنجاح</p>
                                    </td></tr>

                                    <tr><td dir=""ltr"">
                                    <p style='text-align:center;'><span style='font-weight:bold;'>The First Warning Letter due to Low Cumulative Grade Point Average (CGPA) For Bachelor</span></p>
                                    <p>Dear student,</p>

                                    Name: " + student.Profile.Name + @"<br />
                                    Reference No.: " + student.MatrixNo + @"<br />
                                    Date: " + DateTime.Now.ToString("dd MMMM yyyy") + @"<br />

                                    <p>Assalamu’alaikum,</p>
                                    <p>
                                    Due to low CGPA which is less than 2.00 for the first time, Examination & Record Department in Admission & Registration Division would like to inform you 
                                    that you should improve your effort and performance in order to raise your CGPA to become above than 2.00 in your next semester. 
                                    However, you should be aware that our department going to terminate your study if you fail to do so. 
                                    Therefore, please be informed that your termination status will not allow you to further study in our university unless you reapply as a new student.
                                    </p>
                                    <p>Best wishes for your success and excellence.</p>
                                    </td></tr></table>";


            if (svcMessaging.SendEmail(student.Profile.Email, subject, body, true))
            {
                SaveEmailMessage(student, HttpContext.Current.User.Identity.Name, "Low CGPA Warning Email" + "|" + semesterCode, body, "LOW CGPA WARNING", "LOW CGPA", "lowcgpa.html");
                return true;
            }
            return false;
        }

        public bool SendFirstWarningLowCgpaPostgraduate(Student student, string semesterCode)
        {
            string nameAr = String.IsNullOrEmpty(student.Profile.NameAr) ? student.Profile.Name : student.Profile.NameAr;
            var subject = "The First Warning Letter due to Low Cumulative Grade Point Average (CGPA)";
            var body = @"
                            <table>
                            <tr>
                                <td class = ""header"" style=""text-align: center"">
                                     <p><a href=""www.mediu.edu.my""><img src=""http://cms.mediu.edu.my/office/content/logo_color.jpg"" /></a></p> <br />
                                </td>
                            </tr>
                            <tr style='padding-bottom:10px;'><td dir=""rtl"">
                            <p style='text-align:center;'><span style='font-weight:bold;'>خطاب الإنذار الأول بسبب انخفاض المعدل التراكمي - الدراسات العليا</span></p>
                            <p>عزيزي الطالب/ عزيزي الطالبة؛</p>
                            <p>
                                        الإسم : " + nameAr + @"<br />
                                        الرقم المرجعي : " + student.MatrixNo + @" <br />
                                        التاريخ : <span dir=""ltr"">" + DateTime.Now.ToString("dd MMMM yyyy") + @"</span><br />
                            </p>
                            <p>السلام عليكم ورحمة الله وبركاته</p>
                            <p>نظرا لانخفاض معدلكم التراكمي عن أقل نسبة للنجاح (3.00) للمرة الأولى؛ فإن قسم الامتحانات والسجلات بعمادة القبول والتسجيل يوجه لكم هذا الخطاب تنبيها وحثا لكم على ضرورة بذل الجهد والطاقة في رفع معدلكم التراكمي ليصبح (3.00) فما فوق في الفصل التالي، ويتأسف القسم في إلغاء قيدكم (فصلكم) من الدراسة في حالة عجزكم عن تحقيق ذلك، علما بأنه في حالة الفصل لا يمكنكم الدراسة بالجامعة مرة أخرى إلا بتقديم طلب جديد.</p>
                            <p>مع التمنيات لكم بالتوفيق والنجاح</p>
                            </td></tr>

                            <tr><td dir=""ltr"">
                            <p style='text-align:center;'><span style='font-weight:bold;'>The First Warning Letter due to Low Cumulative Grade Point Average (CGPA) For Postgraduate</span></p>
                            <p>Dear student,</p>

                            Name: " + student.Profile.Name + @"<br />
                            Reference No.: " + student.MatrixNo + @"<br />
                            Date: " + DateTime.Now.ToString("dd MMMM yyyy") + @"<br />

                            <p>Assalamu’alaikum,</p>
                            <p>
                            Due to low CGPA which is less than 3.00 for the first time, Examination & Record Department in Admission & Registration Division would like to inform you that you should improve your effort and performance in order to raise your CGPA to become above than 3.00 in your next semester. However, you should be aware that our department going to terminate your study if you fail to do so. Therefore, please be informed that your termination status will not allow you to further study in our university unless you reapply as a new student.
                            </p>
                            <p>Best wishes for your success and excellence.</p>
                            </td></tr></table>";


            if (svcMessaging.SendEmail(student.Profile.Email, subject, body, true))
            {
                SaveEmailMessage(student, HttpContext.Current.User.Identity.Name, "Low CGPA Warning Email" + "|" + semesterCode, body, "LOW CGPA WARNING", "LOW CGPA", "lowcgpa.html");
                return true;
            }
            return false;
        }

        private string DescriptionForStudentAcademicEmail(SubjectRegistered sr)
        {
            return sr.SemesterEvent.Code + "|" + sr.SubjectCode + "|" + sr.ClassSectionName + "|" + sr.TutorialGroupName;
        }

        private void SaveEmailMessage(PersonRole receiver, string sender, string description, string body, string category, string title, string filename)
        {
            var emailDocument = new Mediu.Cms.Domain.Profiles.EmailDocument();
            emailDocument.Data = CreateDocumentForEmail(receiver.Profile.Email, sender, title, body);
            emailDocument.MimeType = "text/html";
            emailDocument.IsSoftCopyAvailable = true;
            emailDocument.Title = title;
            emailDocument.Description = description;
            emailDocument.Name = filename;
            emailDocument.Category = category;
            receiver.Profile.AddEmailDocument(emailDocument);
            emailDocumentSvc.SaveDocument(receiver.Profile, emailDocument);
        }

        private byte[] CreateDocumentForEmail(string receiver, string sender, string subject, string body)
        {
            string combine = "<html>" + "\n" + "<head>" + "\n";
            combine += @"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" + "\n" + "</head>";
            combine += "To: " + receiver + "<br />";
            combine += "From: " + sender + "<br />";
            combine += "DateTime: " + DateTime.Now.ToString() + "<br />";
            combine += "Subject: " + subject + "<br />";
            combine += body;
            combine += "</html>";
            return Encoding.UTF8.GetBytes(combine);
        }

        public bool SendMeetingMinutesNotificationEmail(PostGraduateThesis postgraduateThesis, string refNo, string tutorName, string meetingDate, string meetingTime, string emailSubject, string tutorEmail)
        {

            string body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"">
                                <p>Dear " + tutorName + @"  god bless you</p>
                                <p>السلام عليكم ورحمة الله وبركاته</p>
                                <p>
                                We would like to inform you that the below student has submitted 
                                his proposal to your distinguished department for the thesis defense procedures.
                                </p>
                                <p>
                                Therefore we hope you , -as a mentor or faculty member in this department, 
                                or head of the department- to access into our CMS system and read the proposal, 
                                and send your comments to the head of department within 3 days.
                                </p>

                                <p>Your cooperation is highly appreciated.</p>

                                <p> Student Name :  " + postgraduateThesis.Student.Profile.Name + @" </p>
                                <p> Student Matrix No :  " + postgraduateThesis.Student.MatrixNo + @" </p>

                                <p>
                                Signed by:
                                Scientific Management & Graduation
                                DEANSHIP OF POSTGRADUATE STUDIES
                                Al-madinah International University
                                </p>

                                <p><font color=red>Note: 
                                This e-message was sent  to all faculty members in this department, Accordingly, 
                                please ignore this message if tour not member in this department.</font>
                                </p>
                                <p>
                                    PS: Please send the signed document to ahmed.nour@mediu.edu.my
                                </p>
                                
                            </td>
                                <td width=""50%"" dir=""rtl"">
                                <p> سعادة  : " + tutorName + @" وفقه الله </p>
                                <p>السلام عليكم ورحمة الله وبركاته </p>
                                <p>
                                نحيطكم علمًا بأنّ الطالب المذكور بياناته في أدناه قد رفع خطته إلى قسمكم الموقر عبر نظام (CMS)؛  استعدادا لعرضها في جلسة ""الدفاع عن الخطة"".   
                                </p>

                                <p>
                                وعليه نأمل منكم: بصفتكم مرشدًا، أو عضو هيئة التدريس في هذا القسم، أو رئيس القسم الدخول إلى نظام (CMS)، وقراءة الخطة، ومن ثم إبداء المرئيات وإرسالها إلى رئيس القسم خلال ثلاثة أيام.
                                </p>

                                <p>شاكرين لكم حسن التعاون</p>
                              
                                <p>   اسم الطالب:  " + postgraduateThesis.Student.Profile.NameAr + @" </p>
                                <p>   رقمه المرجعي: " + postgraduateThesis.Student.MatrixNo + @" </p><br/>

                                <p> توقيع
                                الإدارة العلمية والتخرج
                                عمادة الدراسات العليا
                                جامعة المدينة العالمية
                                </p><br/>

                                <p><font color=red>ملحوظة: هذه الرسالة إلكترونية لجميع أعضاء هيئة التدريس المسجلين في القسم الذي ينتمي إليه الطالب؛ وعليه يرجى شاكرًا تجاهلها في حال أنَّكم لستم عضوًا في ذلك القسم.
                                </font></p>   
                                 <p>
                                    ملاحظة: يرجى ارسال الوثيقة موقعة الى 
                                    ahmad.nour@mediu.edu.my
                                </p>       
                            </td>
                        </tr>
                    </table>
                    ";


            return svcMessaging.SendEmail(tutorEmail, emailSubject, body, true);
        }

        public bool SentLetterOfAuditBalance(Student student, string auditBalanceAmount)
        {

            var subject = "LETTER OF BALANCE CONFIRMATION FOR AUDIT PURPOSES";
            var body = @"
                    <table width=""100%"">
                        <tr>
                            <td width=""50%"" dir=""ltr"" valign=""top"">

                            Date: " + DateTime.Now.ToString("dd/MMMM/yyyy") + @"<br />
                            Student Name: " + student.Profile.Name + @"<br />
                            Matriculation No: " + student.MatrixNo + @"<br /><br />
                            
                            Dear Sir/Madam<br /><br />
                                                        
                            <p align =""justify"">
                            With reference to the above, please find enclosed herewith the Confirmation Letter which is self-explanatory.</p>
                            <p align =""justify"">
                            Kindly reply to this e-mail to state your confirmation on the balance as per the letter to this email address: : finance.division@mediu.edu.my</p>
    
                            Thank you for your prompt reply.<br />                       
                            <br /><br />
                            Yours truly<br />
                            for Al-Madinah International University Sdn. Bhd.</br><br /><br /><br />
                            </td>
                            <td dir=""rtl"">
                             <br/>   سيدي العزيز / سيدتي،
                             <p>   الموضوع: تأكيد على صحة الرصيد الغير مسدد لغاية 31/ ديسمبر/ 2015م
                                بالإشارة الى ما سبق ، ولغرض التدقيق المالي ، فإننا نقدر اذا كنت تستطيع تأكيد ما اذا كان الرصيد الغير مسدد من حسابك في سجلاتنا في 31 ديسمبر 2015 هو RM " + auditBalanceAmount + @" صحيح.
                              </p>
                            <p>
                                يرجي الرد علينا بالتأكيد على عنوان البريد الإلكتروني الاتي: finance.division@mediu.edu.my .
                            </p><br/> 
                                هذا ليس طلب للدفع، ولكن طلب للحصول على تأكيد رصيدك.
                            <br/> 
                                شكراً لك.
                            <br/> <br/> 
                                تفضلوا بقبول فائق الاحترام
                                لجامعة المدينة العالمية شبكة التنمية المستدامة. بي اتش دي 

                            </td>
                        </tr>
                    </table>";


            string fileName = "auditbalance_" + student.MatrixNo + ".pdf";


            string fileAuditPath = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\Audit\\" + fileName;

            bool fileCreated = GenerateAuditBalanceLetter(fileAuditPath, student, auditBalanceAmount);

            if (fileCreated)
            {
                try
                {
                    //Send it by Email
                    System.IO.BinaryReader br = new System.IO.BinaryReader(System.IO.File.Open(fileAuditPath, System.IO.FileMode.Open, System.IO.FileAccess.Read));
                    br.BaseStream.Position = 0;
                    byte[] buffer = br.ReadBytes(Convert.ToInt32(br.BaseStream.Length));
                    br.Close();

                    //apply.Student.Profile.Email
                    bool result = svcMessaging.SendEmailWithAttachment(student.Profile.Email, subject, body, true, buffer, fileName);

                    //Take a copy for reviewing after this
                    string copyPath = @"c:\cms_temp\OfferLetter\";

                    string savefilename = Guid.NewGuid().ToString();
                    string copyFile = copyPath + savefilename + ".pdf";
                    if (!Directory.Exists(copyPath))
                    {
                        Directory.CreateDirectory(copyPath);
                    }
                    File.Copy(fileAuditPath, copyFile, true);

                    // Save the pdf file in filestore
                    var doc = new Mediu.Cms.Domain.Profiles.Document();
                    doc.Data = buffer;
                    doc.Category = "AUDITBALANCE";
                    doc.Name = fileName;
                    doc.IsSoftCopyAvailable = true;
                    doc.Title = "Audit balance letter ";
                    doc.Description = doc.Title + " for " + student.MatrixNo;
                    doc.MimeType = "application/pdf";

                    svcDoc.SaveDocument(student.Profile, doc);

                    return result;
                }
                catch
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }

        public bool GenerateAuditBalanceLetter(string destAudit, Student student, string amount)
        {
            DateTime currentDateTime = DateTime.Now;

            var s = sm.OpenSession();
            string sourceDoc;

            sourceDoc = HttpContext.Current.Request.PhysicalApplicationPath + "\\Content\\Docs\\Audit\\ConfirmationAuditBalance.pdf";

            try
            {

                PdfReader r = new PdfReader(sourceDoc);
                string guid = Guid.NewGuid().ToString();
                PdfStamper stamper = new PdfStamper(r, new FileStream(destAudit, FileMode.OpenOrCreate));
                PdfContentByte canvas = stamper.GetOverContent(1);
                BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\arialuni.ttf", BaseFont.IDENTITY_H, true);

                canvas.SetFontAndSize(bf, 12);

                Font fn = new Font(bf, 12, Font.NORMAL, BaseColor.BLACK);
                Font fb = new Font(bf, 12, Font.BOLD, BaseColor.BLACK);
                Font fi = new Font(bf, 12, Font.ITALIC, BaseColor.BLACK);


                Student stu = student;


                // Date
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(currentDateTime.ToString("dd.MMMM.yyyy"), fn), 72, 698, 0);

                //Student name and matrix
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_LEFT, new Phrase(stu.Profile.Name + " (" + stu.MatrixNo + ")", fb), 72, 673, 0);


                //amound 
                ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_LEFT, new Phrase(amount, fb), 385, 507, 0);

                //---------------------------------------

                ////Student name and matrix
                ColumnText.ShowTextAligned(canvas,
                      Element.ALIGN_CENTER, new Phrase(stu.Profile.Name + " (" + stu.MatrixNo + ")", fb), 290, 224, 0);

                ////amound 
                ColumnText.ShowTextAligned(canvas,
                        Element.ALIGN_LEFT, new Phrase(amount, fb), 256, 174, 0);

                stamper.Close();

                return true;

            }
            catch (Exception ex)
            {
                return false;

            }
        }

        public bool SendFailToSendToAX(string custInvoiceTableRecid, string studentAccountNum, string revenueCategoryCode, string invoiceAmount, string currencyCode, string exchRate, string description, string department, string facultyCode, string learningCenterCode)
        {

            var subject = "Invoice is not sent to AX - " + custInvoiceTableRecid;
            var body = @" <table>
                            <tr> The invoice with the details below has problem and not sent to AX" + @"<br />";
            body += "Invoice id: " + custInvoiceTableRecid + @"<br />";
            body += "Rtudent Account Num: " + studentAccountNum + @"<br />";
            body += "Revenue Category Code" + revenueCategoryCode + @"<br />";
            body += "Invoice Amount" + invoiceAmount + @"<br /> </tr></table>";


            if (svcMessaging.SendEmail("zuraida.othman@mediu.edu.my", subject, body, true))
            {
                svcMessaging.SendEmail("ahmad.ahmad@mediu.edu.my", subject, body, true);
                return true;
            }
            return false;
        }

        public bool SendEmailWarningLowCountOfHsbcVirtualAccountNumber(int count)
        {
            string[] recipients = new string[] { "zuraida.othman@mediu.edu.my", "ahmad.ahmad@mediu.edu.my", "zainul.azhan@mediu.edu.my", "raja.roslawati@mediu.edu.my" };
            var subject = "WARNING: Hsbc Virtual Account Number balance in CMS is low";
            var body = @"The total Hsbc Virtual Account Number not yet assign to applicant is " + count + ".";
            body += "Please request new Hsbc Virtual Account Number from HSBC.";
            bool result = false;
            foreach (var item in recipients)
            {
                result = svcMessaging.SendEmail(item, subject, body, true);
            }
            return result;
        }

        public bool SendEmailAvBrochureInEn(AdmissionApply apply)
        {
            var entries = courseRequireRepo.FindAll().Where(c => c.Course.Id == apply.Course1.Id);
            if (apply.LearningMode == LearningMode.OnCampus)
            {
                entries = entries.Where(c => c.EntryRequirementItem.Category == Category.OnCampus);
            }
            else
            {
                entries = entries.Where(c => c.EntryRequirementItem.Category != Category.OnCampus);
            }

            string emailBody = @"
<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">
<html xmlns=""http://www.w3.org/1999/xhtml"" >
<head>
    <title>Welcome to Mediu</title>
    <style type=""text/css"">
    body {position: relative; min-height: 100%; top: 0px;}
    p{font-family:""Calibri"", sans-serif; font-size:12pt; padding: 0px 5px;}
    .title1 {font-family:""Arial"", sans-serif; font-size:14pt; color:white; padding: 5px 5px; }
    .title2 {font-family:""Calibri"", sans-serif; font-size:14pt; color:white; padding: 5px 5px; }
    #sub-header{background-color:#33B7B4;}
    #main-content{
      width:800px;position: absolute;
      margin: auto;
      top: 10px;
      right: 0;
      bottom: 0;
      left: 0;min-width:800px;
      }
    </style>
</head>
<body>
<div id=""main-content"">
<table cellspacing=""0"" style=""background-color:#EEEEEE; width:800px; border-spacing:0px; padding:0px;"">
<tr>
<td>
<a href=""http://www.mediu.edu.my/"">
<img alt=""mediu logo"" width=""800px"" height=""350px"" src=""http://online.mediu.edu.my/apply/Content/Images/logo_en.jpg"" /></a>

<p>Dear Applicant  " + apply.Applicant.Profile.Name + @",<br />
" + "RefNo: " + apply.RefNo + @"

</p>
<p>
<span style=""font-family:Tahoma, Geneva, sans-serif;color:#2F8299;font-size:14pt;"">Assalammualaikum</span>
</p>
<p>
It is our pleasure that you want to join us in " + apply.LearningMode.ToString() + @" study.
</p>
<p> 
And we are pleased to inform you that we have received your application. Now you can login to the <a href=""http://online.mediu.edu.my/applicantportal/Home/Login"">Applicant Portal</a>
through the university site <a href=""http://www.mediu.edu.my/"">www.mediu.edu.my</a> where you can choose Prospect Student using your user name and password which has been emailed to you to follow up your application.
</p>
<p>
If you have any inquiry regarding your application, or you have faced any problem in your application procedures, please feel free to contact our support team <a href=""http://whoson.mediu.edu.my/chat/chatstart.aspx?domain=www.mediu.edu.my"">Via Live Chat</a>
</p>
<p align=""center"" style=""font-weight:bold;font-size:14pt;"">
Following are the information about the program that you want to join:
</p>
<table style=""width:800px"">
    <tr style=""background-color:#33B7B4;"">
    <td style=""width:50%;font-family:Arial, sans-serif; font-size:14pt; color:white;""><span>" + apply.Course1.NameEn + @"</span></td>
    <td align=""center"" style=""width:50%;font-family:Arial, sans-serif; font-size:14pt; color:white;""><span>" + apply.LearningMode.ToString() + @"</span></td>
    </tr>
</table>
<br />
<table style=""width:800px"">
    <tr style=""background-color:#33B7B4;"">
    <td align=""center"" style=""width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>Program Admission Requirements</span></td>
    <td align=""center"" style=""width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>Language Requirements</span></td>
    </tr>
    <tr>
    <td rowspan=""5"" style=""width:50%; vertical-align:top"">" + GetEntryRequirementEnInHtml(entries.ToList()) + @"</td>
        <td style=""width:50%"">
            <p>
            <strong>First:</strong> for applicants who apply for programs that teach in Arabic language and their native language is not Arabic they have to pass MEDIU Arabic placement test.<br />
            <strong>Second:</strong>  for applicants who apply for programs that teaches in English and their native language is not English they have to fulfill one of the following:
            </p>
            <ul>
                <li>Provide proof of achieving TOEFL.</li>
                <li>Provide proof of achieving IELTS.</li>
                <li>Pass MEDIU English placement Test.</li>
            </ul>
        </td>
    </tr>
    <tr>
    <td align=""center"" style=""background-color:#33B7B4; width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>Program Structure</span></td>
    </tr>
    <tr><td align=""center"" style=""width:50%""><a href=""http://online.mediu.edu.my/apply/CoursePlan.aspx?RefNo=" + apply.RefNo + @""">Click here.</a></td></tr>
    <tr>
    <td align=""center"" style=""background-color:#33B7B4; width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>Program Tuition Fees / Program Duration</span></td>
    </tr>
    <tr><td align=""center"" style=""width:50%""><a href=""http://online.mediu.edu.my/apply/CourseFeeStructure.aspx?RefNo=" + apply.RefNo + @""">Click here.</a></td></tr>
</table>
<br />
</td>
</tr>
</table>
<br />
<table style=""width:800px"">
    <tr>
        <td align=""center"" style=""width:40%""><span style=""font-family:'Times New Roman', serif; font-size:12pt; color:#00487E;"">Student Recruitment Unit</span></td>
        <td align=""center"" rowspan=""2"" style=""width:20%"">
        	<a href=""https://twitter.com/mediuuniversity"">
           <img alt=""tweeter"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/tweeter.jpg"" style=""vertical-align:middle"" /></a>
    <a href=""http://www.facebook.com/pages/Al-Madinah-International-University-MEDIU/347978814275"">
           <img alt=""facebook"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/facebook.jpg"" style=""vertical-align:middle""  /></a>
    <a href=""http://www.youtube.com/user/mediuvideo"">
           <img alt=""youtube"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/youtube.jpg"" style=""vertical-align:middle""  /></a>
        </td>
        <td align=""center"" rowspan=""2"" style=""width:40%"">
        <span style=""font-weight:bold;font-size:12pt;"">+603 55113939 </span>
		  <img alt=""phone"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/phone.jpg"" style=""vertical-align:middle"" />
		  </td>
    </tr>
    <tr>
        <td align=""center""><span style=""font-family:'Century Gothic', sans-serif; font-size:12pt; color:#00A6A2;"">Al Madinah International Univesity</span></td>
    </tr>
</table>
<img alt=""footer"" width=""800px"" height=""30px"" src=""http://online.mediu.edu.my/apply/Content/Images/footer.png"" />
</div>
</body>
</html>";

            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, "Welcome to Mediu", emailBody, true);
        }

        public bool SendEmailAvBrochureInAr(AdmissionApply apply)
        {
            var entries = courseRequireRepo.FindAll().Where(c => c.Course.Id == apply.Course1.Id);
            if (apply.LearningMode == LearningMode.OnCampus)
            {
                entries = entries.Where(c => c.EntryRequirementItem.Category == Category.OnCampus);
            }
            else
            {
                entries = entries.Where(c => c.EntryRequirementItem.Category != Category.OnCampus);
            }

            string name = String.IsNullOrEmpty(apply.Applicant.Profile.NameAr) ? apply.Applicant.Profile.Name : apply.Applicant.Profile.NameAr;
            string mode = apply.LearningMode == LearningMode.OnCampus ? "التعليم المباشر" : "التعليم عن بعد";

            string emailBody = @"
<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">
<html xmlns=""http://www.w3.org/1999/xhtml"" >
<head>
    <title>Welcome to Mediu</title>
    <style type=""text/css"">
    body {position: relative; min-height: 100%; top: 0px;}
    p{font-family:""Traditional Arabic"", sans-serif; font-size:14pt; padding: 0px 5px;}
    .title1 {font-family:""Arial"", sans-serif; font-size:14pt; color:white; padding: 5px 5px; }
    .title2 {font-family:""Calibri"", sans-serif; font-size:14pt; color:white; padding: 5px 5px; }
    #sub-header{background-color:#33B7B4;}
    #main-content{
      width:800px;position: absolute;
      margin: auto;
      top: 10px;
      right: 0;
      bottom: 0;
      left: 0;min-width:800px;
      }
    </style>
</head>
<body dir=""rtl"">
<div id=""main-content"">
<table cellspacing=""0px""  style=""background-color:#EEEEEE; width:800px; border-spacing:0px; padding:0px;"">
<tr>
<td>

<a href=""http://www.mediu.edu.my/?lang=ar"">
<img alt=""mediu logo"" width=""800px"" height=""320px"" src=""http://online.mediu.edu.my/apply/Content/Images/logo_ar.jpg"" /></a>

<p>	عزيزي المتقدم " + name + @",<br />
الرقم المرجعي  " + apply.RefNo + @"
</p>
<p>
<span style=""font-family:Tahoma, Geneva, sans-serif;color:#2F8299;font-size:14pt;"">السلام عليكم ورحمة الله وبركاته</span>
</p>
<p>
يسرنا اهتمامكم ورغبتكم في الانضمام إلى ركبنا بنظام " + mode + @"
</p>
<p>
ويسعدنا أن نبلغكم بأن طلبكم قد وصلنا وتم اعتماده، وبإمكانكم الدخول الآن إلى <a href=""http://online.mediu.edu.my/applicantportal/Home/Login"">بوابة المتقدمين</a>  
أو عن طريق موقع الجامعة <a href=""http://www.mediu.edu.my/?lang=ar"">www.mediu.edu.my</a> واختيار بوابة المتقدمين، ومن ثم استخدام اسم المستخدم وكلمة المرور 
التي تم إرساله إليكم، لمتابعة إجراءات طلبكم.
</p>
<p>
وإذا كان لديكم أي استفسار بخصوص طلبكم، أو واجهتكم مشاكل في الإجراءات؛ الرجاء الاتصال بفريق الدعم
 الخاص بالجامعة عبر خدمة <a href=""http://whoson.mediu.edu.my/chat/chatstart.aspx?domain=www.mediu.edu.my"">المحادثة المباشرة</a>
</p>
<p align=""center"" style=""font-weight:bold;font-size:14pt;"">
وفيما يلي معلومات البرنامج الذي ترغبون الالتحاق به
</p>
<table style=""width:800px"">
    <tr style=""background-color:#33B7B4;"">
    <td style=""width:50%;font-family:Arial, sans-serif; font-size:14pt; color:white;""><span>" + apply.Course1.NameAr + @"</span></td>
    <td align=""center"" style=""width:50%;font-family:Arial, sans-serif; font-size:14pt; color:white;""><span>" + mode + @"</span></td>
    </tr>
</table>
<br />
<table style=""width:800px"">
    <tr style=""background-color:#33B7B4;"">
    <td align=""center"" style=""width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>شروط الالتحاق بالبرنامج الدراسي</span></td>
    <td align=""center"" style=""width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>المتطلبات اللغوية</span></td>
    </tr>
    <tr>
    <td rowspan=""5"" style=""width:50%; vertical-align:top"">" + GetEntryRequirementArInHtml(entries.ToList()) + @"</td>
    <td style=""width:50%"">
        <p>
        أولاً: للمتقدم في برامج دراسية تُدرس باللغة العربية وهي ليست لغته الأم فينبغي اجتياز امتحان اللغة العربية الخاص بجامعة المدينة العالمية.<br />
    ثانيًا: للمتقدم في برامح تُدرس باللغة الإنجليزية وهي ليست لغته الأم، يتوجب على الطالب إحراز واحداً مما يأتي: 
        </p>
        <ul>
            <li>تقديم شهادة اجتياز امتحان TOEFL اختبار اللغة الإنجليزية كلغة أجنبية.</li>
            <li>أو تقديم شهادة اجتياز امتحان IELTS خدمة اختبار اللغة الإنجليزية كلغة عالمية.</li>
            <li>أو اجتياز اختبار تحديد مستوى مهارة اللغة الإنجليزية في جامعة المدينة العالمية.</li>
        </ul>
    </td>
    </tr>
    <tr>
    <td align=""center"" style=""background-color:#33B7B4; width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>هيكل الدراسة</span></td>
    </tr>
    <tr><td align=""center"" style=""width:50%""><a href=""http://online.mediu.edu.my/apply/CoursePlan.aspx?RefNo=" + apply.RefNo + @""">إضغط هنا</a></td></tr>
    <tr>
    <td align=""center"" style=""background-color:#33B7B4; width:50%;font-family:Calibri, sans-serif; font-size:14pt; color:white;""><span>رسوم البرنامج / مدة الدراسة</span></td>
    </tr>
    <tr><td align=""center"" style=""width:50%""><a href=""http://online.mediu.edu.my/apply/CourseFeeStructure.aspx?RefNo=" + apply.RefNo + @""">إضغط هنا</a></td></tr>
</table>
<br />
</td>
</tr>
</table>

<br />
<table style=""width:800px"" dir=""ltr"">
    <tr>
        <td align=""center"" style=""width:40%""><span style=""font-family:'Times New Roman', serif; font-size:12pt; color:#00487E;"">وحدة استقطاب الطلاب</span></td>
        <td align=""center"" rowspan=""2"" style=""width:20%"">
        	<a href=""https://twitter.com/mediuuniversity"">
           <img alt=""tweeter"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/tweeter.jpg"" style=""vertical-align:middle"" /></a>
    <a href=""http://www.facebook.com/pages/Al-Madinah-International-University-MEDIU/347978814275"">
           <img alt=""facebook"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/facebook.jpg"" style=""vertical-align:middle""  /></a>
    <a href=""http://www.youtube.com/user/mediuvideo"">
           <img alt=""youtube"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/youtube.jpg"" style=""vertical-align:middle""  /></a>
        </td>
        <td align=""center"" rowspan=""2"" style=""width:40%"">
        <span style=""font-weight:bold;font-size:12pt;"">+603 55113939 </span>
		  <img alt=""phone"" width=""34px"" height=""34px"" src=""http://online.mediu.edu.my/apply/Content/Images/phone.jpg"" style=""vertical-align:middle"" />
		  </td>
    </tr>
    <tr>
        <td align=""center""><span style=""font-family:'Century Gothic', sans-serif; font-size:12pt; color:#00A6A2;"">جامعة المدينة العالمية</span></td>
    </tr>
</table>
<img alt=""footer"" width=""800px"" height=""30px"" src=""http://online.mediu.edu.my/apply/Content/Images/footer.png"" />
</div>
</body>
</html>
";

            return svcMessaging.SendEmail(apply.Applicant.Profile.Email, "أهلاً بكم في جامعة المدينة العالمية", emailBody, true);
        }

        public bool SendEmailContactUsInAr(string email)
        {
            string body = @"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">
<html xmlns=""http://www.w3.org/1999/xhtml"" >
<head>
    <title></title>
    <style  type=""text/css"">
    body {position: relative; min-height: 100%; top: 0px;}
    #main-content{
      width:80%;position: absolute;
      margin: auto;
      top: 10px;
      right: 0;
      bottom: 0;
      left: 0;min-width:800px;
    }
    .link_header
    {
    	padding-bottom: 40px;
    }
    .link_row
    {
    	padding-bottom: 30px;
    }
    </style>
</head>
<body>
<div id=""main-content"">
<table>
<tr><td class=""link_header"" align=""right""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/header.jpg""/></td></tr>
<tr><td class=""link_row"" align=""right""><a href= ""http://www.mediu.edu.my/contact-us-ar/?lang=ar""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/link1.jpg""/></a></td></tr>
<tr><td class=""link_row"" align=""right""><a href= ""mailto:admission@mediu.edu.my""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/link2.jpg""/></a></td></tr>
<tr><td class=""link_row"" align=""right""><a href= ""http://whoson.mediu.edu.my/chat/chatstart.aspx?domain=www.mediu.edu.my&session=310-1397029099994&SID=0""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/link3.jpg""/></a></td></tr>
</table>
</div>
</body>
</html>
";
            return svcMessaging.SendEmail(email, "للتواصل معنا", body, true);
        }

        public bool SendEmailContactUsInEn(string email)
        {
            string body = @"
            <!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">
            <html xmlns=""http://www.w3.org/1999/xhtml"" >
            <head>
                <title></title>
                <style  type=""text/css"">
                    #main-content{
                  width:80%;position: absolute;
                  margin: auto;
                  top: 10px;
                  right: 0;
                  bottom: 0;
                  left: 0;min-width:800px;
                }
                .link_header
                {
	                padding-bottom: 40px;
                }
                .link_row
                {
	                padding-bottom: 30px;
                }
                </style>
            </head>
            <body>
            <div id=""main-content"">
            <table>
            <tr><td class=""link_header""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/header_en.jpg""/></td></tr>
            <tr><td class=""link_row""><a href= ""http://www.mediu.edu.my/support/contact-us-2/""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/link1_en.jpg""/></a></td></tr>
            <tr><td class=""link_row""><a href= ""mailto:admission@mediu.edu.my""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/link2_en.jpg""/></a></td></tr>
            <tr><td class=""link_row""><a href= ""http://whoson.mediu.edu.my/chat/chatstart.aspx?domain=www.mediu.edu.my&session=310-1397029099994&SID=0""><img width=""800px"" height=""auto"" src=""http://online.mediu.edu.my/apply/Content/Images/Brochure/link3_en.jpg""/></a></td></tr>
            </table>
            </div>
            </body>
            </html>";
            return svcMessaging.SendEmail(email, "Contact Us", body, true);

        }

        public void SendEmailOnlineMeetingAnnouncementToAllAttendee(Meeting meeting, bool isChange)
        {
            foreach (var attendee in meeting.Attendees)
            {
                try
                {
                    string linkEn = "";
                    string linkAr = "";
                    string dear = "سعادة ";
                    switch (attendee.MeetingAttendeeType)
                    {
                        case MeetingAttendeeType.Staff:
                            linkEn = "You can go to the meeting from CMS Office.";
                            linkAr = "يمكنك حضور الاجتماع (اللقاء) عبر نظام الحرم الجامعي CMS";
                            break;
                        case MeetingAttendeeType.Lecturer:
                            linkEn = "You can go to the meeting from Academic Portal.";
                            linkAr = "يمكنك حضور الاجتماع (اللقاء) عبر البوابة الأكاديمية";
                            break;
                        case MeetingAttendeeType.Student:
                            linkEn = "You can go to the meeting from Student Portal.";
                            linkAr = "يمكنك حضور الاجتماع (اللقاء) عبر بوابة الطالب";
                            dear = "عزيزي ";
                            break;
                        case MeetingAttendeeType.External:
                            linkEn = @"You can go to the meeting from this <a href= ""http://academicportal.mediu.edu.my/postgraduate/gotomeetingexternal?key=" + attendee.MeetingKey + @""">link</a>";
                            linkAr = @"يمكنك حضور الاجتماع (اللقاء) عبر" + @"<a href= ""http://academicportal.mediu.edu.my/postgraduate/gotomeetingexternal?key=" + attendee.MeetingKey + @""">الرابط المباشر</a>";
                            break;
                    }
                    var body = @"
            <table width=""100%"">
                <tr>
                    <td colspan=""2""><div><center><img src=""https://cms.mediu.edu.my/office/content/logo_color.jpg"" alt=""mediu"" /></center></div></td>
                </tr>
                <tr>
                    <td width=""50%"" dir=""ltr"">
                    <p>Dear " + attendee.Name + @",</p>
		            <p>السلام عليكم ورحمة الله وبركاته</p>
                    <p>SUBJECT: Online Meeting.</p>";
                    if (isChange)
                    {
                        body += @"<p>Please be informed that there are changes in the Date or Time for the following online meeting.</p>
		            <p>";
                    }
                    else
                    {
                        body += @"<p>Please be informed that the following online meeting has been created and you are invited.</p>
		            <p>";
                    }

		            body += @"<b>Meeting Title :</b>" + meeting.Name + @"<br />
		            <b>Start Date    : </b>" + meeting.StartDate.ToString("dd MMMM yyyy") + @" <br />
		            <b>Start Time    :</b>" + meeting.DisplayStartTime + " (Malaysia Time)" + @" <br />
		            <b>End Time  :</b>" + meeting.DisplayEndTime + " (Malaysia Time)" + @" <br />
                    <b>Presenter :</b>" + meeting.MeetingPresenter + @" <br />
		            </p>
		            <p>" + linkEn + @"</p>
		            <p>Thank you</p>
		            <br />
		            <p>This email is computer generated and no signature required.</p>
		            </td>

		            <td width=""50%"" dir=""rtl"">
                    <p>" + dear + attendee.Name + @",</p>
                    <p> الموضوع: اللقاءات / الاجتماعات.</p>
                    <p>السلام عليكم ورحمة الله وبركاته</p>";
                    if (isChange)
                    {
                        body += @"<p>يرجى العلم أن هناك تغييرات في تاريخ ووقت القاءات الاونلاين التالية:</p>
		            <p>";
                    }
                    else
                    {
                        body += @"
                    <p>يرجى العلم أنّ اللقاء الافتراضي أدناه قد تمّ إنشاؤه وأنك مدعو للحضور وفقا للموعد أدناه: </p>
                    <p>";
                    }
                    body += @"
                    <b>عنوان الاجتماع (اللقاء):</b>" + meeting.GetNameArabic + @"<br/>
                    <b>تاريخ الاجتماع (اللفاء):</b>" + meeting.StartDate.ToString("dd MMMM yyyy") + @"<br/>
                    <b>موعد بدأ اللقاء:</b>" + meeting.DisplayStartTime + "-" + "بتوقيت ماليزيا" + @"<br/>
                    <b>موعد انتهاء اللقاء:</b>" + meeting.DisplayEndTime + "-" + "بتوقيت ماليزيا" + @"<br/>
                    <b>مقدّم اللقاء:</b>" + meeting.MeetingPresenter + @" <br /><br/>
                    </p>
                    <p>" + linkAr + @"</p>
                    <p>وشكرا لكم.</p>
                    <br /> 
                    <p>هذه الرسالة الكترونية من أجل الإشعار ولا تتطلب التوقيع أو الرد. </p>
	
		            </td>
	            </tr>
            </table>";
                    svcMessaging.SendEmail(attendee.Email, "Online Meeting", body, true);
                }
                catch (Exception err)
                {
                }
            }
        }

        private string GetEntryRequirementEnInHtml(IList<CourseEntryRequirement> courseEntryRequirements)
        {
            string htmlContent = "<ul>";
            foreach (var item in courseEntryRequirements)
            {
                htmlContent += "<li>" + item.EntryRequirementItem.DescriptionEn + "</li>";
            }
            htmlContent += "</ul>";
            return htmlContent;
        }

        private string GetEntryRequirementArInHtml(IList<CourseEntryRequirement> courseEntryRequirements)
        {
            string htmlContent = "<ul>";
            foreach (var item in courseEntryRequirements)
            {
                htmlContent += "<li>" + item.EntryRequirementItem.DescriptionAr + "</li>";
            }
            htmlContent += "</ul>";
            return htmlContent;
        }

    }

}