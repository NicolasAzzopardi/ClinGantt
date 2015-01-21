rm(list=ls())
if(grepl("nicolas",getwd())){setwd("C:/Users/Nicolas/Documents/Gantt")}
setwd("C:/Users/Nicolas/git/CliGantt")
toInstall <- c("RODBC","ggplot2","gridExtra","grid","scales","plyr")
doInstall=NULL;for(p in toInstall){doInstall=c(doInstall,!any(row.names(installed.packages())==p))}
if(any(doInstall)){install.packages(toInstall[doInstall], repos = "http://cran.r-project.org")}
lapply(toInstall, library, character.only = TRUE)
#################################################################################

File = "141001_Gantt"

xlsConnect <- odbcConnectExcel(paste(File,".xls",sep=""))
	CLI=sqlFetch(xlsConnect, "Etudes_Cliniques")
	ANI=sqlFetch(xlsConnect, "Etudes_Animales")
	COH=as.data.frame(sqlFetch(xlsConnect, "Etudes_Cohortes"))
	EVT=sqlFetch(xlsConnect, "Evènement")[,1:5]
	DOS=sqlFetch(xlsConnect, "Dosage")
	INC=sqlFetch(xlsConnect, "Inclusion-Suivi")
	MOD=sqlFetch(xlsConnect, "Modélisation")
	PUB=sqlFetch(xlsConnect, "Publication")
odbcClose(xlsConnect)

# Suppr les études non plannifiées
CLI = CLI[CLI$ETAT!="Contact",]
CLI = CLI[CLI$ETAT!="Non retenu",]
CLI = CLI[CLI$ETAT!="Pas de participations",]
CLI = CLI[CLI$ETAT!="En attente",]

# Suppr de certaines études
CLI =  CLI[CLI$ETUDE!="Humeur Aqueuse",]
CLI =  CLI[CLI$ETUDE!="DICV (cohort DEFI)",]
CLI =  CLI[CLI$ETUDE!="SMART",]

INC = INC[,c(1:4,6)]
INC = INC[!is.na(INC$ETUDE),]
INC = INC[!is.na(INC$DATE_debut),]
INC = INC[!is.na(INC$DATE_fin),]
INC = with(INC,data.frame(ETUDE,DATE_debut,DATE_fin,Qui= "Autre",Quoi,Commentaires,Etape= "Données"))

CLI$Type="Clinique"
ANI$Type="Animale"
COH$Type="Cohorte"

CLI = CLI[!is.na(CLI$ETUDE),]
ANI = ANI[!is.na(ANI$ETUDE),]
COH = COH[!is.na(COH$ETUDE),]

CLI$Num = as.numeric(1:dim(CLI)[1])
ANI$Num = as.numeric(1:dim(ANI)[1])
COH$Num = as.numeric(1:dim(COH)[1])

ListeEtudes = rbind.fill(CLI,COH,ANI)
ListeEtudes = data.frame(ETUDE = ListeEtudes$ETUDE
						,Type = ListeEtudes$Type
						,Num = ListeEtudes$Num
						)

names(EVT)[grep("DATE",names(EVT))] = "DATE_debut"
EVT$DATE_fin = NA

EVT$Etape[EVT$Quoi=="tubes"] = "Dosage"
EVT$Etape[EVT$Quoi=="1er contact"] = "Données"

MOD$Etape = "Modélisation"
PUB$Etape = "Publication"

DOS = with(DOS,data.frame(ETUDE,DATE_debut,DATE_fin,Qui,Quoi="Dosage",Commentaires,Etape="Dosage"))


DOS$DATE_fin[DOS$Commentaires=="En cours"&!is.na(DOS$Commentaires)] = as.Date(Sys.Date())

DosPrevEtudes = DOS$ETUDE[is.na(DOS$DATE_debut)]
for(etu in DosPrevEtudes){
DOS$DATE_debut[DOS$ETUDE==etu] = max(INC$DATE_fin[INC$ETUDE==etu&INC$Quoi=="Suivi PK"])
}

DosPrevEtudes = DOS$ETUDE[is.na(DOS$DATE_debut)]
for(etu in DosPrevEtudes){
DOS$DATE_debut[DOS$ETUDE==etu] = max(INC$DATE_fin[INC$ETUDE==etu&INC$Quoi=="Inclusion"],na.rm=TRUE)
}

DOS$DATE_fin[is.na(DOS$DATE_fin)] = as.Date(DOS$DATE_debut[is.na(DOS$DATE_fin)])+3*365.25/12




meanPKTime = as.numeric(mean(with(MOD[as.Date(MOD$DATE_fin)+1<as.Date(Sys.Date())&MOD$Quoi=="PK"&(MOD$Qui=="David"|MOD$Qui=="Nicolas"),],DATE_fin-DATE_debut),na.rm=TRUE))/60/60/24#
meanPKPDTime = as.numeric(mean(with(MOD[as.Date(MOD$DATE_fin)+1<as.Date(Sys.Date())&MOD$Quoi=="PKPD"&(MOD$Qui=="David"|MOD$Qui=="Nicolas"),],DATE_fin-DATE_debut),na.rm=TRUE))#/60/60/24#

ModPKPrevEtudes = MOD$ETUDE[is.na(MOD$DATE_debut)&MOD$Quoi=="PK"]
for(etu in ModPKPrevEtudes){
MOD$DATE_debut[MOD$ETUDE==etu&MOD$Quoi=="PK"] = max(DOS$DATE_fin[DOS$ETUDE==etu])
}
MOD$DATE_fin[is.na(MOD$DATE_fin)&!is.na(MOD$DATE_debut)] = as.Date(MOD$DATE_debut[is.na(MOD$DATE_fin)&!is.na(MOD$DATE_debut)])+meanPKTime#3*365.25/12

ModPKPDPrevEtudes = MOD$ETUDE[is.na(MOD$DATE_debut)&MOD$Quoi=="PKPD"]
for(etu in ModPKPDPrevEtudes){
MOD$DATE_debut[MOD$ETUDE==etu&MOD$Quoi=="PKPD"] = MOD$DATE_fin[MOD$ETUDE==etu&MOD$Quoi=="PK"]
}

MOD$DATE_fin[is.na(MOD$DATE_fin)] = as.Date(MOD$DATE_debut[is.na(MOD$DATE_fin)])+meanPKPDTime#7*365.25/12



meanPUBTime = as.numeric(mean(with(PUB[as.Date(PUB$DATE_fin)+1<as.Date(Sys.Date())&(PUB$Qui=="David"|PUB$Qui=="Nicolas"),],DATE_fin-DATE_debut),na.rm=TRUE))#/60/60/24#


PubPrevEtudes = PUB$ETUDE[is.na(PUB$DATE_debut)]
for(etu in PubPrevEtudes){
PUB$DATE_debut[PUB$ETUDE==etu] = max(MOD$DATE_fin[MOD$ETUDE==etu])
}
PUB$DATE_fin[is.na(PUB$DATE_fin)] = as.Date(PUB$DATE_debut[is.na(PUB$DATE_fin)])+meanPUBTime#8*365.25/12











DATA = rbind(INC,DOS,MOD,EVT,PUB)
DATA = DATA[!is.na(DATA$ETUDE),]


DATA$Etape <- factor(DATA$Etape, levels = c("Données","Dosage","Modélisation","Publication"))
DATA$NumEtape = as.numeric(DATA$Etape)

DATA$ETUDE <- factor(DATA$ETUDE, levels = ListeEtudes$ETUDE)

DATA = merge(DATA,ListeEtudes)

names(DATA) =  c("ETUDE", "starts", "ends", "Qui", "Evénement","COM", "Etape","NumEtape","Type","Num" )
DATA$starts = as.Date(DATA$starts) +1 
DATA$ends = as.Date(DATA$ends) +1 

DATA = DATA[!is.na(DATA$starts),]
DATA = arrange(DATA,starts)

DATA$ETUDE <- factor(DATA$ETUDE, levels = unique(DATA$ETUDE))





ThemeGantt = theme(	axis.text.x =  element_text(colour = "black",hjust=1, vjust = .5,angle = 90),
		axis.text.y =  element_text(size = 0, lineheight = 0.9, colour = "transparent"),
		axis.ticks.y = element_line(colour = "transparent"),
		axis.ticks.x = element_line(colour = "#00000010", size = 0.1),

		panel.background =  element_rect(fill = "transparent", colour = NA), 
		panel.border =      element_blank(), 
		panel.grid.major.y =  element_line(colour = "transparent"),
		panel.grid.minor.y =  element_line(colour = "transparent"),
		panel.grid.major.x =  element_line(colour = "#00000040", size = 0.1),
		panel.grid.minor.x =  element_line(colour = "#00000010", size = 0.05),
		panel.margin =      unit(0, "lines"),

		legend.title =      element_text(size =5, face = "bold", hjust = 0),
		legend.text  = element_text(size =5),
		#legend.position = c(0, 0),
		legend.justification = c(0,0),
		#legend.position = "bottom", legend.box = "horizontal",
		strip.text.y = element_text(size =5, hjust = 0, vjust = 0.5,angle = 0),
		strip.background =  element_rect(fill = "#00000020", colour = NA),

		plot.title =        element_text(hjust = 0),
		plot.margin =       unit(c(1,1,1,1), "lines")
	)

ThemeGanttCharge = theme(	axis.text.x =  element_text(colour = "black",hjust=1, vjust = .5,angle = 90),
		# axis.text.y =  element_text(size = 0, lineheight = 0.9, colour = "transparent", hjust = 1),
		# axis.ticks.y = element_line(colour = "transparent"),
		axis.ticks.x = element_line(colour = "#00000040", size = 0.1),

		panel.background =  element_rect(fill = "transparent", colour = NA), 
		panel.border =      element_blank(), 
		panel.grid.major.y =  element_line(colour = "#00000010", size = 0.05),
		panel.grid.minor.y =  element_line(colour = "transparent"),
		panel.grid.major.x =  element_line(colour = "#00000040", size = 0.1),
		panel.grid.minor.x =  element_line(colour = "#00000010", size = 0.05),
		panel.margin =      unit(0, "lines"),

		legend.title =      element_text(size =5, face = "bold", hjust = 0),
		legend.text  = element_text(size =5),
		legend.position = c(1, 1),
		legend.justification = c(1, 1),
		#legend.position = "bottom", legend.box = "horizontal",
		strip.text.y = element_text(size =4, hjust = 0, vjust = 0.5,angle = 0),
		strip.background =  element_rect(fill = "#00000010", colour = NA),

		plot.title =        element_text(hjust = 0),
		plot.margin =       unit(c(1,1,1,1), "lines")
	)

plotGantt <- function(DATA,func,Qui="Global",Range_T=NA){
	if(any(is.na(Range_T))){
		Xlim = c(as.Date(min(format(DATA$starts,"%Y-01-01"))),as.Date(max(format(c(DATA$starts,DATA$ends),"%Y-12-31"),na.rm=TRUE)))
		Break_T = "years"
		Name_T = date_format("%Y")
	} else {
		Xlim = c(as.Date(Sys.Date()+Range_T[1]*365.2525/12),as.Date(Sys.Date()+Range_T[2]*365.2525/12))
		if(diff(Range_T)>25){
			Break_T = "years"
			Name_T = date_format("%Y")
		}else{
			Break_T = "months"
			Name_T = date_format("%m-%Y")
		}
	}
	if(Qui!="Global") DATA = DATA[DATA$Qui==Qui,]
DATA = arrange(DATA,starts)
DATA$ends[is.na(DATA$ends)] = DATA$starts[is.na(DATA$ends)]
DATA = DATA[DATA$ends>Xlim[1]&DATA$starts<Xlim[2],]

DATA$ETUDE <- factor(DATA$ETUDE, levels = unique(DATA$ETUDE))

ggplot() +
	geom_rect(data=DATA[!is.na(DATA$ends),]
		,aes(xmin=as.Date(starts), xmax=as.Date(ends), ymin=+1-NumEtape/4, ymax=+1-(NumEtape-1)/4, fill=Etape)
		,xlim=Xlim,ylim=c(0,1), alpha=0.8
	)+ 
	geom_point(data=DATA[is.na(DATA$ends),]
	,aes(x=starts,y=NumEtape,shape=Evénement)
	,size=.5,col="red",xlim=Xlim,ylim=c(0,1)
	)+
	geom_vline(xintercept = as.numeric(as.Date(Sys.Date()),xlim=Xlim),col="blue")+ 
	facet_grid(func,shrink = TRUE,drop=TRUE)+
	labs(title=paste("Diagramme de Gantt des études au ",format(Sys.Date(),format="%d %b %Y")," (",Qui,")",sep="")
		,x = "",y="")+
	guides(fill = guide_legend(title = "Etape"))+
	# scale_x_date(labels=date_format("%b %Y"), breaks="6 month", minor_breaks = "1 month",limits=Xlim)+
	scale_x_date(labels=Name_T, breaks=seq.Date(Xlim[1],Xlim[2],by=Break_T),limits=Xlim)+
	expand_limits(colour = factor(c("Données","Dosage")))+

	ThemeGantt#+ coord_fixed(ratio=30)
}


ploGlobal_6m <- plotGantt(DATA,"ETUDE~.","Global",c(0,6))
ploDavid_6m <- plotGantt(DATA,"ETUDE~.","David",c(0,6))
ploNicolas_6m <- plotGantt(DATA,"ETUDE+Type~.","Nicolas",c(0,6))
ploCLIGlobal_6m <- plotGantt(DATA[DATA$Type=="Clinique",],"ETUDE~.","Global",c(0,6))
ploANIGlobal_6m <- plotGantt(DATA[DATA$Type=="Animale",],"ETUDE~.","Global",c(0,6))
ploCOHGlobal_6m <- plotGantt(DATA[DATA$Type=="Cohorte",],"ETUDE~.","Global",c(0,6))
ggsave(paste("./6_mois/",File,"_GanttGlobal_6m.jpg",sep=""), plot=ploGlobal_6m,width=18,units="cm")
ggsave(paste("./6_mois/",File,"_GanttDavid_6m.jpg",sep=""), plot=ploDavid_6m,width=18,height=10,units="cm")
ggsave(paste("./6_mois/",File,"_GanttNicolas_6m.jpg",sep=""), plot=ploNicolas_6m,width=18,height=10,units="cm")
ggsave(paste("./6_mois/",File,"_GanttCLI_6m.jpg",sep=""), plot=ploCLIGlobal_6m,width=18,height=10,units="cm")
ggsave(paste("./6_mois/",File,"_GanttANI_6m.jpg",sep=""), plot=ploANIGlobal_6m,width=18,height=10,units="cm")
ggsave(paste("./6_mois/",File,"_GanttCOH_6m.jpg",sep=""), plot=ploCOHGlobal_6m,width=18,height=10,units="cm")

ploGlobal_18m <- plotGantt(DATA,"ETUDE~.","Global",c(0,18))
ploDavid_18m <- plotGantt(DATA,"ETUDE~.","David",c(0,18))
ploNicolas_18m <- plotGantt(DATA,"ETUDE+Type~.","Nicolas",c(0,18))
ploCLIGlobal_18m <- plotGantt(DATA[DATA$Type=="Clinique",],"ETUDE~.","Global",c(0,18))
ploANIGlobal_18m <- plotGantt(DATA[DATA$Type=="Animale",],"ETUDE~.","Global",c(0,18))
ploCOHGlobal_18m <- plotGantt(DATA[DATA$Type=="Cohorte",],"ETUDE~.","Global",c(0,18))
ggsave(paste("./18_mois/",File,"_GanttGlobal_18m.jpg",sep=""), plot=ploGlobal_18m,width=18,units="cm")
ggsave(paste("./18_mois/",File,"_GanttDavid_18m.jpg",sep=""), plot=ploDavid_18m,width=18,height=10,units="cm")
ggsave(paste("./18_mois/",File,"_GanttNicolas_18m.jpg",sep=""), plot=ploNicolas_18m,width=18,height=10,units="cm")
ggsave(paste("./18_mois/",File,"_GanttCLI_18m.jpg",sep=""), plot=ploCLIGlobal_18m,width=18,units="cm")
ggsave(paste("./18_mois/",File,"_GanttANI_18m.jpg",sep=""), plot=ploANIGlobal_18m,width=18,height=10,units="cm")
ggsave(paste("./18_mois/",File,"_GanttCOH_18m.jpg",sep=""), plot=ploCOHGlobal_18m,width=18,height=10,units="cm")

ploGlobal <- plotGantt(DATA,"ETUDE~.","Global")
ploDavid <- plotGantt(DATA,"ETUDE~.","David")
ploNicolas <- plotGantt(DATA,"ETUDE+Type~.","Nicolas")
ploCLIGlobal <- plotGantt(DATA[DATA$Type=="Clinique",],"ETUDE~.","Global")
ploANIGlobal <- plotGantt(DATA[DATA$Type=="Animale",],"ETUDE~.","Global")
ploCOHGlobal <- plotGantt(DATA[DATA$Type=="Cohorte",],"ETUDE~.","Global")
ggsave(paste("./tout/",File,"_GanttGlobal.jpg",sep=""), plot=ploGlobal,width=18,units="cm")
ggsave(paste("./tout/",File,"_GanttDavid.jpg",sep=""), plot=ploDavid,width=18,height=10,units="cm")
ggsave(paste("./tout/",File,"_GanttNicolas.jpg",sep=""), plot=ploNicolas,width=18,height=10,units="cm")
ggsave(paste("./tout/",File,"_GanttCLI.jpg",sep=""), plot=ploCLIGlobal,width=18,height=10,units="cm")
ggsave(paste("./tout/",File,"_GanttANI.jpg",sep=""), plot=ploANIGlobal,width=18,height=10,units="cm")
ggsave(paste("./tout/",File,"_GanttCOH.jpg",sep=""), plot=ploCOHGlobal,width=18,height=10,units="cm")



MODDTNA = MOD[MOD$Qui=="David"|MOD$Qui=="Nicolas",]
MODDavid = MOD[MOD$Qui=="David",]
MODNicolas = MOD[MOD$Qui=="Nicolas",]

PUBDTNA = PUB[PUB$Qui=="David"|PUB$Qui=="Nicolas",]
PUBDavid = PUB[PUB$Qui=="David",]
PUBNicolas = PUB[PUB$Qui=="Nicolas",]

Cumul  <-function(DAT){
Dat = rbind( data.frame(date=DAT$DATE_debut,val=1) , data.frame(date=DAT$DATE_fin,val=-1))
Dat = Dat[order(Dat$date),]
Dat$Cumu = NA
Dat$Cumu[1] = 1
for(li in 2:dim(Dat)[1]){Dat$Cumu[li] = Dat$Cumu[li-1] + Dat$val[li]}
return(Dat[,-2])
}

ChargeMod = Cumul(MODDTNA)
ChargeMod$Activité = "Modélisation"
ChargeModDavid = Cumul(MODDavid)
ChargeModDavid$Activité = "Modélisation"
ChargeModNicolas = Cumul(MODNicolas)
ChargeModNicolas$Activité = "Modélisation"

ChargePub = Cumul(PUBDTNA)
ChargePub$Activité = "Publication"
ChargePubDavid = Cumul(PUBDavid)
ChargePubDavid$Activité = "Publication"
ChargePubNicolas = Cumul(PUBNicolas)
ChargePubNicolas$Activité = "Publication"

Charge = rbind(ChargeMod,ChargePub)
ChargeDavid = rbind(ChargeModDavid,ChargePubDavid)
ChargeNicolas = rbind(ChargeModNicolas,ChargePubNicolas)


PlotCharge <- function(DATA){
DATA$date = as.Date(DATA$date)

Xlim = c(as.Date(min(format(DATA$date,"%Y-01-01"))),as.Date(max(format(DATA$date,"%Y-12-31"),na.rm=TRUE)))
Ylims = c(0,max(DATA$Cumu)+2)
ggplot(DATA,aes(x=date,Cumu,group=Activité,col=Activité)) +
	geom_step(data=DATA[(as.Date(DATA$date)-as.Date(Sys.Date())<0),],linetype="solid")+
	geom_step(data=DATA,linetype="F1")+
	geom_vline(xintercept = as.numeric(as.Date(Sys.Date()),xlim=Xlim),col="blue")+ 
	scale_y_continuous(breaks=0:max(Ylims),limits=Ylims)+
	scale_x_date(labels=date_format("%Y"), breaks=seq.Date(Xlim[1],Xlim[2],by="year"),limits=Xlim)+
	labs(title="",x = "",y="Nombre cumulé")+
	expand_limits(colour = factor(c("Données","Dosage")))+
	ThemeGanttCharge
}

PloCharge = PlotCharge(Charge)
PloChargeDavid  = PlotCharge(ChargeDavid )
PloChargeNicolas = PlotCharge(ChargeNicolas)

pdf(file=paste(File,".pdf",sep=""),paper="a4")
ploGlobal
grid.arrange(ploGlobal, PloCharge , ncol=1,main="Global")
grid.arrange(ploDavid, PloChargeDavid, ncol=1,main="David")
grid.arrange(ploNicolas, PloChargeNicolas, ncol=1,main="Nicolas")
dev.off()

ggplot(data=DATA	,aes(x=ends,y=as.numeric(ends-starts)/365.25*12,col=Etape))+geom_point()+ stat_smooth()

# ggplot(DATA[DATA$Qui=="David"|DATA$Qui=="Nicolas",],aes(x=as.numeric(ends-starts)/365.25*12,fill=Etape))+ geom_density(alpha=.5)+
# facet_grid(Qui~.,shrink = FALSE)


# ggplot(DATA[DATA$Qui=="David"|DATA$Qui=="Nicolas",],aes(x=as.numeric(ends-starts)/365.25*12,fill=Qui))+ geom_density(alpha=.5)+
# facet_grid(Etape~.,shrink = FALSE)
ggsave(paste("./charge/",File,"_ChargeGlobale.jpg",sep=""), plot=PloCharge,width=18,height=10,units="cm")
ggsave(paste("./charge/",File,"_ChargeDavid.jpg",sep=""), plot=PloChargeDavid,width=18,height=10,units="cm")
ggsave(paste("./charge/",File,"_ChargeNicolas.jpg",sep=""), plot=PloChargeNicolas,width=18,height=10,units="cm")



