# RestAPI-Scripts


#***********************************************************************************************#
#                                                                                               #
#    Title       : Phishing Email Deletion using Azure Compliance Search                        #
#    Author      : Udit Mahajan                                                                 #
#    Date        : 03/06/2023                                                                   #
#    Code version: V3                                                                           #
#                                                                                               #
#***********************************************************************************************#

#***********************************************************************************************#
#                                                                                               #
#    Pre-requisites:                                                                            #
#       1. Ensure Execution policy is set to remotely signed                                    #     
#       2. Azure Compliance Manager or Global Administrator roles                               #
#       3. A connected Exchange Online account                                                  #
#***********************************************************************************************#

1.   The code is able to identify emails based on Date,Subject,Body and Sender email address    
     and delete all emails based on the citeria.                                                
2.   Please check if Content search results are as desired as it is possible to delete all      
     organisation mailboxes accidentally                                                        
3.   I have intentially commented delete option to prevent accidental deletes   
