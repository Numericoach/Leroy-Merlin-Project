import { Utils } from "../utils";

/*
 Testez ici vos fonctionnalités de back-log une par une et insérez les fonctions de test dans testAll() pour les exécuter toutes en même temps.
 Il s'agit ici de tester les fonctionnalités liées à Apps Script et donc aux outils Google et non les fonctions de traitement qui elles se font dans le dossier `local-tests`.
*/

/**
 * Fonction de test de toutes les fonctionnalités de back-log
 */
function testAll(){
    testGetUserEmail();
}

/**
 * Exemple de fonction de test
 */
function testGetUserEmail(){
    try{
        const USER_EMAIL = Utils.getUserEmail();
        /* Traitement si besoin ...*/
        
        console.log("✅ getUserEmail() OK ✅");
    }catch(error){
        console.log(error);
        throw new Error("❌ getUserEmail() K.O ❌")
    }
}